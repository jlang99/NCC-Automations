import os, sys, threading, time
from datetime import datetime, timezone
from twilio.rest import Client

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_dir)
from PythonTools import PHONE, TWILIO_MSGS
# Run this file Standalone to test the Robocall feature. Will Call Joseph and the NCC Desk with the estimated highest probability of occuring. NCC Desk is a recorded line.


# ─── Channels to monitor ──────────────────────────────────────────────────────
MONITORED_CHANNELS = [
    'Q REACTIVE POWER',
    'PF POWER FACTOR',
]

# ─── Change → flow map ────────────────────────────────────────────────────────
# Keys are frozensets of (channel, from_val, to_val) tuples representing the
# exact set of changes detected in a single check.  None means OCR could not
# read the value.
#
#   Single-element frozenset  → fires when ONLY that one channel changed.
#   Multi-element frozenset   → fires when ALL those channels changed together.
#
# Each value must be a key that exists in TWILIO_MSGS (PythonTools.py).
# Add a new row here any time you create a new Studio flow.
CHANGE_FLOW_MAP = {
    # ── Single-channel known transitions ─────────────────────────────────────
    frozenset([('Q REACTIVE POWER', 'ENABLED',  'DISABLED')]): 'Lily PPC Q E->D',
    frozenset([('Q REACTIVE POWER', 'DISABLED', 'ENABLED' )]): 'Lily PPC Q D->E',
    frozenset([('PF POWER FACTOR',  'DISABLED', 'ENABLED' )]): 'Lily PPC PF D->E',
    frozenset([('PF POWER FACTOR',  'ENABLED',  'DISABLED')]): 'Lily PPC PF E->D',

    # ── Both channels change in the same check ────────────────────────────────
    frozenset([                                                # both flip away from default
        ('Q REACTIVE POWER', 'ENABLED',  'DISABLED'),
        ('PF POWER FACTOR',  'DISABLED', 'ENABLED'),
    ]): 'Lily PPC Vars Flipped',
    frozenset([                                                # both return to default
        ('Q REACTIVE POWER', 'DISABLED', 'ENABLED'),
        ('PF POWER FACTOR',  'ENABLED',  'DISABLED'),
    ]): 'Lily PPC Back to Default',

    # ── Variable not found (OCR returned None) ────────────────────────────────
    frozenset([('Q REACTIVE POWER', 'ENABLED',  None)]): 'Lily PPC Q Not Found',
    frozenset([('Q REACTIVE POWER', 'DISABLED', None)]): 'Lily PPC Q Not Found',
    frozenset([('PF POWER FACTOR',  'DISABLED', None)]): 'Lily PPC PF Not Found',
    frozenset([('PF POWER FACTOR',  'ENABLED',  None)]): 'Lily PPC PF Not Found',
    frozenset([                                                # both unreadable in same check
        ('Q REACTIVE POWER', 'ENABLED',  None),
        ('PF POWER FACTOR',  'DISABLED', None),
    ]): 'Lily PPC Both Not Found',
    frozenset([                                                # both unreadable in same check
        ('Q REACTIVE POWER', 'DISABLED',  None),
        ('PF POWER FACTOR',  'ENABLED', None),
    ]): 'Lily PPC Both Not Found',

    # ── Variable recovered from None ─────────────────────────────────────────
    frozenset([('Q REACTIVE POWER', None, 'ENABLED' )]): 'Lily PPC Q Recovered',
    frozenset([('Q REACTIVE POWER', None, 'DISABLED')]): 'Lily PPC Q Recovered Q Disabled',
    frozenset([('PF POWER FACTOR',  None, 'DISABLED')]): 'Lily PPC PF Recovered',
    frozenset([('PF POWER FACTOR',  None, 'ENABLED' )]): 'Lily PPC PF Recovered PF Enabled',
    frozenset([                                                # both readable again in same check
        ('Q REACTIVE POWER', None, 'ENABLED'),
        ('PF POWER FACTOR',  None, 'DISABLED'),
    ]): 'Lily PPC Both Recovered Default',
    frozenset([                                                # both readable again in same check
        ('Q REACTIVE POWER', None, 'DISABLED'),
        ('PF POWER FACTOR',  None, 'ENABLED'),
    ]): 'Lily PPC Both Recovered Flipped',
}

# Who receives the robocall on any PPC alert
CALL_RECIPIENTS = [
    PHONE['Joseph Lang'], PHONE['NCC']
]

# ─── Retry settings ───────────────────────────────────────────────────────────
RETRY_INTERVAL_SEC = 5 * 60   # total time between call attempts
CALL_SETTLE_SEC    = 120      # wait after dialing before checking call status
MAX_RETRIES        = 48       # give up after ~4 hours

# ─── Active retry tracking ────────────────────────────────────────────────────
_active_retries: dict = {}
_lock = threading.Lock()


def _was_answered(client, recipient, after_dt):
    """Return True if a call to recipient placed after after_dt was answered."""
    calls = client.calls.list(to=recipient, start_time_after=after_dt, limit=20)
    return any(c.status in ('completed', 'in-progress') for c in calls)


def _retry_loop(flow_sid, recipient, key):
    """Background thread: dials recipient every RETRY_INTERVAL_SEC until answered."""
    account_sid = os.environ.get('TWILIO_ACCOUNT_SID')
    auth_token  = os.environ.get('TWILIO_AUTH_TOKEN')
    client = Client(account_sid, auth_token)

    for attempt in range(1, MAX_RETRIES + 1):
        call_time = datetime.now(timezone.utc)

        try:
            execution = client.studio.v2.flows(flow_sid).executions.create(
                to=recipient,
                from_=PHONE['Twilio'],
            )
            print(f'Robocall attempt {attempt} -> {recipient} (sid={execution.sid})')
        except Exception as exc:
            print(f'Robocall attempt {attempt} failed to initiate: {exc}')

        time.sleep(CALL_SETTLE_SEC)

        try:
            if _was_answered(client, recipient, call_time):
                print(f'Robocall answered on attempt {attempt} -> {recipient}')
                break
        except Exception as exc:
            print(f'Could not verify call status (attempt {attempt}): {exc}')

        remaining = RETRY_INTERVAL_SEC - CALL_SETTLE_SEC
        if remaining > 0 and attempt < MAX_RETRIES:
            time.sleep(remaining)
    else:
        print(f'Robocall retry limit ({MAX_RETRIES}) reached -> {recipient}')

    with _lock:
        _active_retries.pop(key, None)


# ─── Public function ──────────────────────────────────────────────────────────

def make_calls(previous, current):
    """Fire the appropriate Twilio Studio flow when any monitored channel changes.

    previous: dict from the last saved state file
    current:  dict of just-detected values (None means OCR could not read it)

    Builds a frozenset of (channel, from, to) tuples for every channel that
    changed, then looks it up in CHANGE_FLOW_MAP to select the right flow.
    If the combination is not in the map, the call is skipped and logged.
    Retries every RETRY_INTERVAL_SEC in a background thread until answered.
    """
    if not previous:
        return

    changes = frozenset(
        (ch, previous.get(ch), current.get(ch))
        for ch in MONITORED_CHANNELS
        if previous.get(ch) != current.get(ch)
    )
    if not changes:
        return

    account_sid = os.environ.get('TWILIO_ACCOUNT_SID')
    auth_token  = os.environ.get('TWILIO_AUTH_TOKEN')
    if not account_sid or not auth_token:
        print('TWILIO_ACCOUNT_SID or TWILIO_AUTH_TOKEN not set — skipping robocall.')
        return

    msg_key = CHANGE_FLOW_MAP.get(changes)
    if not msg_key:
        print(f'No CHANGE_FLOW_MAP entry for changes {sorted(changes)} — skipping robocall.')
        return

    flow_sid = TWILIO_MSGS.get(msg_key)
    if not flow_sid:
        print(f'No flow SID in TWILIO_MSGS for key "{msg_key}" — skipping.')
        return

    for ch, prev_val, curr_val in sorted(changes, key=lambda x: x[0]):
        print(f'PPC change: {ch}  {prev_val} -> {curr_val}')
    print(f'Using flow: {msg_key}')

    for recipient in CALL_RECIPIENTS:
        key = ('ppc_alert', recipient)
        with _lock:
            if key in _active_retries and _active_retries[key].is_alive():
                print(f'Retry loop already active for {recipient}')
                continue
            t = threading.Thread(target=_retry_loop, args=(flow_sid, recipient, key), daemon=True)
            _active_retries[key] = t
            t.start()
        print(f'Robocall retry loop started -> {recipient}')


# ─── Standalone test ──────────────────────────────────────────────────────────

if __name__ == '__main__':
    previous = {'Q REACTIVE POWER': 'ENABLED',  'PF POWER FACTOR': 'DISABLED'}
    current  = {'Q REACTIVE POWER': 'DISABLED', 'PF POWER FACTOR': 'DISABLED'}
    make_calls(previous, current)
    time.sleep(RETRY_INTERVAL_SEC * MAX_RETRIES)
