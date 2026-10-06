import re
from core.engine_utils import Rule, SUPS

COORD_RE = re.compile(r"(?<![\w])(\d{1,3})\s*[*oº]\s*(\d{1,2})['’]?\s*([NSWEBDTN])(?![\w])", re.IGNORECASE)

def N(t): return (t, None)

def get_rules():
    R = []
    # Replace 100o 30' N -> 100° 30' N
    R.append(Rule(COORD_RE, lambda mt: [N(mt.group(1) + "° "), N(mt.group(2) + "' "), N(mt.group(3).upper())]))
    return R
