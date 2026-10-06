# -*- coding: utf-8 -*-
import re
from core.engine_utils import Rule, SUPS, SUBS

ELEMENTS = set("H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr".split())
CHEM_BLACKLIST = re.compile(r"^(HK|HS|PC|BC|CB|SC|NC|KS|CN|CS\d{3,}|CV|PB|GV|TB)\d*$")
_GROUP = r"(?:[A-Z][a-z]?\d*|\((?:[A-Z][a-z]?\d*)+\)\d*)+"
CHEM_RE = re.compile(r"(?<![\w])(\d*)(%s(?:[.\·*]\d*%s)?)(\^?\d*[+-])?(?![\w])" % (_GROUP, _GROUP))

def chem_parts(coef, formula, charge):
    if CHEM_BLACKLIST.match(formula): return None
    syms = re.findall(r"[A-Z][a-z]?", formula)
    if not syms or not all(s in ELEMENTS for s in syms): return None
    has_sub = re.search(r"[A-Za-z)]\d", formula)
    if not has_sub and not charge: return None
    
    body, ch_sup = formula, None
    if charge:
        if charge.startswith("^"): ch_sup = charge[1:]
        else:
            sign, dig = charge[-1], charge[:-1]
            if dig: ch_sup = dig + sign
            else:
                tail = re.search(r"(\d+)$", body)
                single = len(syms) == 1 and "(" not in body
                if tail and single:
                    body, ch_sup = body[:-len(tail.group(1))], tail.group(1) + sign
                elif tail and len(tail.group(1)) >= 2:
                    body, ch_sup = body[:-1], tail.group(1)[-1] + sign
                else:
                    ch_sup = sign
    parts = []
    if coef: parts.append((coef, None))
    prev = ""
    for tok in re.finditer(r"[A-Za-z()]+|\d+|[.\·*]", body):
        t = tok.group(0)
        if t.isdigit(): parts.append((t, SUBS if prev and (prev[-1].isalpha() or prev[-1] == ")") else None))
        elif t in ".\·*": parts.append(("·", None))
        else: parts.append((t, None))
        prev = t
    if ch_sup: parts.append((ch_sup, SUPS))
    return parts

def get_rules():
    R = []
    R.append(Rule(re.compile(r"<-{1,2}>|<=>|⇌"), lambda mt: [(("⇌", None))]))
    R.append(Rule(re.compile(r"-{1,2}>"), lambda mt: [(("→", None))]))
    R.append(Rule(CHEM_RE, lambda mt: chem_parts(mt.group(1), mt.group(2), mt.group(3))))
    return R
