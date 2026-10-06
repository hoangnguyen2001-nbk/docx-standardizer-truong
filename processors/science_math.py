# -*- coding: utf-8 -*-
import re
from core.engine_utils import Rule, SUPS, SUBS

# Define proper Unicode characters
GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", 
    "epsilon": "ε", "theta": "θ", "lambda": "λ", "omega": "ω", 
    "phi": "φ", "varphi": "φ", "rho": "ρ", "sigma": "σ", "pi": "π"
}

def N(t): return (t, None)

def get_rules():
    R = []
    
    # 1. Logic & Relation Arrows
    R.append(Rule(re.compile(r"<-{1,2}>|<=>"), lambda mt: [N("⇔" if mt.group(0) == "<=>" else "↔")]))
    R.append(Rule(re.compile(r"-{1,2}>"), lambda mt: [N("→")]))
    R.append(Rule(re.compile(r"=>"), lambda mt: [N("⇒")]))
    R.append(Rule(re.compile(r">="), lambda mt: [N("≥")]))
    R.append(Rule(re.compile(r"<="), lambda mt: [N("≤")]))
    R.append(Rule(re.compile(r"!=|=/="), lambda mt: [N("≠")]))
    R.append(Rule(re.compile(r"\+/-|\+-(?=\s*\d|\s*[a-zA-Z(])"), lambda mt: [N("±")]))
    
    # 2. Area and Volume Units
    R.append(Rule(
        re.compile(r"(?<![\w])(km|m|cm|dm|mm|ha)(2|3)(?![\w])", re.IGNORECASE),
        lambda mt: [N(mt.group(1)), (mt.group(2), SUPS)]
    ))
    
    # 3. Physics & Math Speed / Complex Units
    R.append(Rule(
        re.compile(r"(?<![\w])(m/s|km/h|rad/s|g/cm|kg/m|g/m|mol/m|N/m)([23])(?![\w])"),
        lambda mt: [N(mt.group(1)), (mt.group(2), SUPS)]
    ))
    
    # 4. Temperature °C
    R.append(Rule(re.compile(r"(?<=\d)(?:\s?\^[0o°]|\s?[o°]| 0)(?=\s?C(?!\w))"), lambda mt: [N("°")]))
    R.append(Rule(re.compile(r"(?<=\d)(?:\^[0o°]|°|o)(?![\w])"), lambda mt: [N("°")]))
    
    # 5. Caret ^ and Underscore _ notation
    R.append(Rule(
        re.compile(r"(?<=[^\s^_])\^(?:\{([^{}]+)\}|\(([^()]+)\)|([+-]?[0-9A-Za-z]+(?:[+-](?![\w(]))?))"),
        lambda mt: [(mt.group(1) or mt.group(2) or mt.group(3), SUPS)]
    ))
    R.append(Rule(
        re.compile(r"(?<=[^\s^_])_(?:\{([^{}]+)\}|([0-9]+|[A-Za-z](?![\w])))"),
        lambda mt: [(mt.group(1) or mt.group(2), SUBS)]
    ))
    
    # 6. Variables with sub-indices
    R.append(Rule(
        re.compile(r"(?<![\w])(T|t|v|x|y|a|F|P|E|I|U)(max|min)(?![\w])"),
        lambda mt: [N(mt.group(1)), (mt.group(2), SUBS)]
    ))
    R.append(Rule(
        re.compile(r"(?<![\w])([A-Za-zαβγδεθλωφρσπ])(\d{1,2})(?![\w])"),
        lambda mt: [N(mt.group(1)), (mt.group(2), SUBS)]
    ))
    
    # 7. Greek letters
    R.append(Rule(
        re.compile(r"(?<![\w])(%s)(?![\w])" % "|".join(sorted(GREEK, key=len, reverse=True))),
        lambda mt: [N(GREEK[mt.group(1)])]
    ))
    
    return R
