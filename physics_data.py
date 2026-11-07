"""Physics data for the formula worksheet generator.

Holds
  * the unit tables (every unit the app understands, with its factor to the SI unit)
  * the GCSE physics equation sheet (as sympy expressions, so any equation can be
    rearranged and evaluated exactly)
  * helpers for number formatting and for safely evaluating an arithmetic string

To add an equation, add one `Equation(...)` entry to EQUATIONS at the bottom.
"""

import ast
import math
import operator
import re
from fractions import Fraction as Fr

import sympy as sp

# --------------------------------------------------------------------------- #
# Units
# --------------------------------------------------------------------------- #

_K = [2, 5, 3.5, 12, 0.4, 7.5]            # examples for kilo-/mega-style units
_MILLI = [250, 40, 1500, 800, 60, 2500]   # examples for milli-style units
_CENTI = [60, 250, 45, 8, 120, 30]        # examples for centi-style units


def _group(si, *others):
    """(SI unit symbol, [(symbol, factor_to_SI, [example values]), ...])"""
    return {"si": si, "others": list(others)}


# dimension -> unit group.  Factors are exact fractions.
UNIT_GROUPS = {
    "force": _group("N", ("kN", Fr(1000), _K), ("MN", Fr(10**6), [2, 0.5, 1.5]), ("mN", Fr(1, 1000), _MILLI)),
    "length": _group("m", ("cm", Fr(1, 100), _CENTI), ("mm", Fr(1, 1000), _MILLI), ("km", Fr(1000), _K)),
    "energy": _group("J", ("kJ", Fr(1000), _K), ("MJ", Fr(10**6), [2, 0.5, 1.5, 3.5])),
    "time": _group("s", ("ms", Fr(1, 1000), _MILLI), ("min", Fr(60), [2, 5, 10, 0.5, 3]), ("h", Fr(3600), [2, 0.5, 1.5, 3, 0.25])),
    "mass": _group("kg", ("g", Fr(1, 1000), [500, 250, 40, 1200, 75]), ("tonnes", Fr(1000), [2, 0.5, 3.5, 12])),
    "speed": _group("m/s", ("km/h", Fr(5, 18), [36, 72, 18, 54, 90]), ("cm/s", Fr(1, 100), _CENTI)),
    "accel": _group("m/s²"),
    "power": _group("W", ("kW", Fr(1000), _K), ("MW", Fr(10**6), [2, 0.5, 1.5, 3.5]), ("mW", Fr(1, 1000), _MILLI)),
    "pressure": _group("Pa", ("kPa", Fr(1000), _K), ("MPa", Fr(10**6), [2, 0.5, 1.5]), ("N/cm²", Fr(10**4), [2, 5, 3.5, 12])),
    "area": _group("m²", ("cm²", Fr(1, 10**4), [200, 5000, 350, 1500]), ("mm²", Fr(1, 10**6), [500, 2500, 8000, 40])),
    "volume": _group("m³", ("cm³", Fr(1, 10**6), [500, 2500, 40000, 150000])),
    "density": _group("kg/m³", ("g/cm³", Fr(1000), [2, 0.8, 1.5, 7.8, 11])),
    "voltage": _group("V", ("mV", Fr(1, 1000), _MILLI), ("kV", Fr(1000), _K)),
    "current": _group("A", ("mA", Fr(1, 1000), _MILLI)),
    "resistance": _group("Ω", ("kΩ", Fr(1000), _K), ("MΩ", Fr(10**6), [2, 0.5, 1.5, 3.5])),
    "charge": _group("C", ("mC", Fr(1, 1000), _MILLI), ("μC", Fr(1, 10**6), [500, 40, 250, 1500])),
    "frequency": _group("Hz", ("kHz", Fr(1000), _K), ("MHz", Fr(10**6), [2, 0.5, 1.5, 3.5, 100]), ("GHz", Fr(10**9), [2, 0.5, 1.5, 3.5])),
    "shc": _group("J/kg°C", ("kJ/kg°C", Fr(1000), [2, 4.2, 0.5, 0.9])),
    "latent": _group("J/kg", ("kJ/kg", Fr(1000), [2, 5, 330, 12]), ("MJ/kg", Fr(10**6), [2, 0.5, 2.3])),
    "gfield": _group("N/kg"),
    "springk": _group("N/m", ("N/cm", Fr(100), [2, 5, 3.5, 12]), ("kN/m", Fr(1000), [2, 0.5, 1.5, 3.5])),
    "tempchange": _group("°C"),
    "moment": _group("Nm", ("Ncm", Fr(1, 100), _CENTI), ("kNm", Fr(1000), [2, 0.5, 1.5])),
    "momentum": _group("kg m/s"),
    "fluxdensity": _group("T", ("mT", Fr(1, 1000), _MILLI)),
}

# symbol -> info, for every unit the app understands (SI units included)
UNIT_INFO = {}
for _dim, _g in UNIT_GROUPS.items():
    UNIT_INFO[_g["si"]] = {"dim": _dim, "factor": Fr(1), "si": _g["si"], "examples": []}
    for _sym, _factor, _examples in _g["others"]:
        UNIT_INFO[_sym] = {"dim": _dim, "factor": _factor, "si": _g["si"], "examples": _examples}

# Longest symbols first so "km/h" is matched before "km" and "m/s²" before "m/s".
_UNIT_ALT = "|".join(re.escape(u) for u in sorted(UNIT_INFO, key=len, reverse=True))
UNIT_IN_TEXT = re.compile(
    r"(?<![\w.,])(\d[\d,]*(?:\.\d+)?)\s*(" + _UNIT_ALT + r")(?![A-Za-z0-9²³°Ω/])"
)
NUMBER_IN_TEXT = re.compile(r"(?<![\w.])\d[\d,]*(?:\.\d+)?")


def normalise_text(text):
    """Tidy characters an LLM often substitutes (micro sign, minus sign, thin spaces)."""
    return (
        str(text)
        .replace("µ", "μ")   # micro sign -> Greek mu
        .replace("−", "-")        # minus sign
        .replace(" ", " ")
        .replace(" ", " ")
    )


def units_in_text(text):
    """Every (number, unit) pair written in a piece of text, e.g. [(20.0, 'cm'), (3.0, 'N')]."""
    found = []
    for match in UNIT_IN_TEXT.finditer(normalise_text(text)):
        found.append((float(match.group(1).replace(",", "")), match.group(2)))
    return found


def numbers_in_text(text):
    return [float(m.group(0).replace(",", "")) for m in NUMBER_IN_TEXT.finditer(normalise_text(text))]


# --------------------------------------------------------------------------- #
# Number formatting and safe arithmetic
# --------------------------------------------------------------------------- #

def fmt_num(value, sig=3):
    """Plain-decimal number to `sig` significant figures, never scientific notation."""
    value = float(value)
    if value == 0:
        return "0"
    rounded = float(f"{value:.{sig}g}")
    text = f"{rounded:.{sig}g}"
    if "e" in text or "E" in text:
        from decimal import Decimal
        text = format(Decimal(text), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def is_exact_to_sig(value, sig=3):
    value = float(value)
    return abs(value - float(f"{value:.{sig}g}")) <= 1e-9 * abs(value)


_OPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.Pow: operator.pow, ast.USub: operator.neg, ast.UAdd: operator.pos,
}


def safe_eval(expression):
    """Evaluate a plain arithmetic string (numbers, + - * / ** ( ) and sqrt())."""
    expression = normalise_text(expression).replace("×", "*").replace("÷", "/").replace("^", "**")
    if len(expression) > 400:
        raise ValueError("calculation_si is too long.")

    def walk(node):
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return float(node.value)
        if isinstance(node, ast.BinOp) and type(node.op) in _OPS:
            left, right = walk(node.left), walk(node.right)
            if isinstance(node.op, ast.Pow) and abs(right) > 10:
                raise ValueError("Exponent too large in calculation_si.")
            return _OPS[type(node.op)](left, right)
        if isinstance(node, ast.UnaryOp) and type(node.op) in _OPS:
            return _OPS[type(node.op)](walk(node.operand))
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "sqrt"
                and len(node.args) == 1 and not node.keywords):
            return math.sqrt(walk(node.args[0]))
        raise ValueError("calculation_si may only contain numbers, + - * / **, brackets and sqrt().")

    try:
        result = walk(ast.parse(expression, mode="eval"))
    except (SyntaxError, ZeroDivisionError, OverflowError) as exc:
        raise ValueError(f"calculation_si could not be evaluated: {exc}") from exc
    if not math.isfinite(result):
        raise ValueError("calculation_si did not give a finite number.")
    return result


# --------------------------------------------------------------------------- #
# Equations
# --------------------------------------------------------------------------- #

_SUPERSCRIPTS = str.maketrans("0123456789-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻")


class Equation:
    """One equation from the equation sheet.

    quantities: list of (key, display symbol, name, dimension).
      * `key` is the ASCII name used inside `expr` (and by the LLM in its JSON)
      * the display symbol may contain <sub> tags, e.g. E<sub>k</sub>
    expr:      "W = F*s"  (left side is the subject of the equation)
    triangle:  (top, bottom_left, bottom_right) keys when a formula triangle is possible
    """

    def __init__(self, name, words, letters, expr, quantities, triangle=None):
        self.name = name
        self.words = words
        self.letters = letters
        self.expr = expr
        self.quantities = quantities
        self.triangle = triangle
        self.keys = [q[0] for q in quantities]
        self.symbol = {q[0]: q[1] for q in quantities}
        self.qname = {q[0]: q[2] for q in quantities}
        self.dim = {q[0]: q[3] for q in quantities}

        self._syms = {k: sp.Symbol(k) for k in self.keys}
        lhs, rhs = (side.strip() for side in expr.split("="))
        self.subject = lhs
        self._solutions = {self.subject: sp.sympify(rhs, locals=self._syms)}
        for key in self.keys:
            if key == self.subject:
                continue
            candidates = sp.solve(sp.Eq(self._syms[self.subject], self._solutions[self.subject]), self._syms[key])
            candidates = [c for c in candidates if not c.could_extract_minus_sign()]
            if len(candidates) != 1:
                raise ValueError(f"Could not rearrange {expr!r} for {key}")
            self._solutions[key] = candidates[0]

    # -- titles / display -------------------------------------------------- #
    @property
    def title(self):
        return f"{self.name} ({self.words[0].upper() + self.words[1:]})"

    def si_unit(self, key):
        return UNIT_GROUPS[self.dim[key]]["si"]

    def allowed_units(self, key):
        group = UNIT_GROUPS[self.dim[key]]
        return [group["si"]] + [u[0] for u in group["others"]]

    # -- maths --------------------------------------------------------------- #
    def evaluate(self, find, values_si):
        """values_si: {key: Fraction}.  Returns the value of `find` as a float."""
        subs = {self._syms[k]: sp.Rational(v.numerator, v.denominator) for k, v in values_si.items()}
        return float(self._solutions[find].subs(subs))

    def rearranged(self, find):
        return f"{self.symbol[find]} = {self._fmt(self._solutions[find])}"

    def substituted(self, find, values_si):
        subs = {self._syms[k]: sp.Symbol(fmt_num(float(v), 6)) for k, v in values_si.items()}
        return self._fmt(self._solutions[find].subs(subs))

    # -- expression -> text (uses x and ÷ rather than sympy's * and /) ---------- #
    def _atom(self, node):
        if isinstance(node, sp.Symbol):
            return self.symbol.get(node.name, node.name)
        if node == sp.Rational(1, 2):
            return "½"
        return str(node)

    def _factor(self, node):
        if isinstance(node, sp.Pow):
            base, exp = node.as_base_exp()
            if exp == sp.Rational(1, 2):
                inner = self._fmt(base)
                return f"√{inner}" if isinstance(base, sp.Symbol) else f"√({inner})"
            if exp.is_Integer and exp > 1:
                return self._atom(base) + str(exp).translate(_SUPERSCRIPTS)
        return self._atom(node)

    def _fmt(self, expr):
        if not isinstance(expr, sp.Mul):
            return self._factor(expr)
        numer, denom = [], []
        coeff = sp.Integer(1)
        args = list(sp.Mul.make_args(expr))
        # sympy splits sqrt(2*E/m) into sqrt(2)*sqrt(E/m); put it back under one root sign
        roots = [a for a in args if isinstance(a, sp.Pow) and a.exp == sp.Rational(1, 2)]
        if len(roots) > 1:
            args = [a for a in args if a not in roots]
            args.append(sp.Pow(sp.Mul(*[r.base for r in roots]), sp.Rational(1, 2), evaluate=False))
        for arg in args:
            if arg.is_Number:
                coeff = coeff * arg
                continue
            base, exp = arg.as_base_exp()
            if exp.is_Rational and exp < 0:
                denom.append(self._factor(base ** (-exp)))
            else:
                numer.append(self._factor(arg))
        p, q = sp.fraction(coeff) if coeff != 1 else (1, 1)
        if coeff == sp.Rational(1, 2) and not denom:
            numer.insert(0, "½")
        else:
            if p != 1 or not numer:
                numer.insert(0, str(p))
            if q != 1:
                denom.insert(0, str(q))
        top = " × ".join(numer)
        if not denom:
            return top
        bottom = denom[0] if len(denom) == 1 else "(" + " × ".join(denom) + ")"
        return f"{top} ÷ {bottom}"


def _eq(name, words, letters, expr, quantities, triangle=None):
    return Equation(name, words, letters, expr, quantities, triangle)


EQUATIONS = {e.name: e for e in [
    _eq("Weight", "weight = mass × gravitational field strength", "W = m × g", "W = m*g",
        [("W", "W", "weight", "force"), ("m", "m", "mass", "mass"), ("g", "g", "gravitational field strength", "gfield")],
        ("W", "m", "g")),
    _eq("Work done", "work done = force × distance", "W = F × s", "W = F*s",
        [("W", "W", "work done", "energy"), ("F", "F", "force", "force"), ("s", "s", "distance", "length")],
        ("W", "F", "s")),
    _eq("Force on a spring", "force = spring constant × extension", "F = k × e", "F = k*e",
        [("F", "F", "force", "force"), ("k", "k", "spring constant", "springk"), ("e", "e", "extension", "length")],
        ("F", "k", "e")),
    _eq("Moment of a force", "moment = force × distance", "M = F × d", "M = F*d",
        [("M", "M", "moment", "moment"), ("F", "F", "force", "force"), ("d", "d", "distance from the pivot", "length")],
        ("M", "F", "d")),
    _eq("Pressure", "pressure = force ÷ area", "P = F ÷ A", "P = F/A",
        [("P", "P", "pressure", "pressure"), ("F", "F", "force", "force"), ("A", "A", "area", "area")],
        ("F", "P", "A")),
    _eq("Distance travelled", "distance = speed × time", "s = v × t", "s = v*t",
        [("s", "s", "distance", "length"), ("v", "v", "speed", "speed"), ("t", "t", "time", "time")],
        ("s", "v", "t")),
    _eq("Acceleration", "acceleration = change in velocity ÷ time", "a = Δv ÷ t", "a = dv/t",
        [("a", "a", "acceleration", "accel"), ("dv", "Δv", "change in velocity", "speed"), ("t", "t", "time", "time")],
        ("dv", "a", "t")),
    _eq("Resultant force", "force = mass × acceleration", "F = m × a", "F = m*a",
        [("F", "F", "force", "force"), ("m", "m", "mass", "mass"), ("a", "a", "acceleration", "accel")],
        ("F", "m", "a")),
    _eq("Momentum", "momentum = mass × velocity", "p = m × v", "p = m*v",
        [("p", "p", "momentum", "momentum"), ("m", "m", "mass", "mass"), ("v", "v", "velocity", "speed")],
        ("p", "m", "v")),
    _eq("Kinetic energy", "kinetic energy = ½ × mass × speed²", "E<sub>k</sub> = ½ × m × v²", "Ek = m*v**2/2",
        [("Ek", "E<sub>k</sub>", "kinetic energy", "energy"), ("m", "m", "mass", "mass"), ("v", "v", "speed", "speed")]),
    _eq("Gravitational potential energy", "gravitational potential energy = mass × gravitational field strength × height",
        "E<sub>p</sub> = m × g × h", "Ep = m*g*h",
        [("Ep", "E<sub>p</sub>", "gravitational potential energy", "energy"), ("m", "m", "mass", "mass"),
         ("g", "g", "gravitational field strength", "gfield"), ("h", "h", "height", "length")]),
    _eq("Elastic potential energy", "elastic potential energy = ½ × spring constant × extension²",
        "E<sub>e</sub> = ½ × k × e²", "Ee = k*e**2/2",
        [("Ee", "E<sub>e</sub>", "elastic potential energy", "energy"), ("k", "k", "spring constant", "springk"),
         ("e", "e", "extension", "length")]),
    _eq("Power (energy transferred)", "power = energy transferred ÷ time", "P = E ÷ t", "P = E/t",
        [("P", "P", "power", "power"), ("E", "E", "energy transferred", "energy"), ("t", "t", "time", "time")],
        ("E", "P", "t")),
    _eq("Power (work done)", "power = work done ÷ time", "P = W ÷ t", "P = W/t",
        [("P", "P", "power", "power"), ("W", "W", "work done", "energy"), ("t", "t", "time", "time")],
        ("W", "P", "t")),
    _eq("Charge flow", "charge flow = current × time", "Q = I × t", "Q = I*t",
        [("Q", "Q", "charge flow", "charge"), ("I", "I", "current", "current"), ("t", "t", "time", "time")],
        ("Q", "I", "t")),
    _eq("Potential difference", "potential difference = current × resistance", "V = I × R", "V = I*R",
        [("V", "V", "potential difference", "voltage"), ("I", "I", "current", "current"), ("R", "R", "resistance", "resistance")],
        ("V", "I", "R")),
    _eq("Electrical power (P = VI)", "power = potential difference × current", "P = V × I", "P = V*I",
        [("P", "P", "power", "power"), ("V", "V", "potential difference", "voltage"), ("I", "I", "current", "current")],
        ("P", "V", "I")),
    _eq("Electrical power (P = I²R)", "power = current² × resistance", "P = I² × R", "P = I**2*R",
        [("P", "P", "power", "power"), ("I", "I", "current", "current"), ("R", "R", "resistance", "resistance")]),
    _eq("Energy transferred (power × time)", "energy transferred = power × time", "E = P × t", "E = P*t",
        [("E", "E", "energy transferred", "energy"), ("P", "P", "power", "power"), ("t", "t", "time", "time")],
        ("E", "P", "t")),
    _eq("Energy transferred (charge × pd)", "energy transferred = charge flow × potential difference", "E = Q × V", "E = Q*V",
        [("E", "E", "energy transferred", "energy"), ("Q", "Q", "charge flow", "charge"),
         ("V", "V", "potential difference", "voltage")],
        ("E", "Q", "V")),
    _eq("Density", "density = mass ÷ volume", "ρ = m ÷ V", "rho = m/V",
        [("rho", "ρ", "density", "density"), ("m", "m", "mass", "mass"), ("V", "V", "volume", "volume")],
        ("m", "rho", "V")),
    _eq("Wave speed", "wave speed = frequency × wavelength", "v = f × λ", "v = f*lam",
        [("v", "v", "wave speed", "speed"), ("f", "f", "frequency", "frequency"), ("lam", "λ", "wavelength", "length")],
        ("v", "f", "lam")),
    _eq("Specific heat capacity", "change in thermal energy = mass × specific heat capacity × temperature change",
        "ΔE = m × c × Δθ", "dE = m*c*dT",
        [("dE", "ΔE", "change in thermal energy", "energy"), ("m", "m", "mass", "mass"),
         ("c", "c", "specific heat capacity", "shc"), ("dT", "Δθ", "temperature change", "tempchange")]),
    _eq("Specific latent heat", "energy for a change of state = mass × specific latent heat", "E = m × L", "E = m*L",
        [("E", "E", "energy for a change of state", "energy"), ("m", "m", "mass", "mass"),
         ("L", "L", "specific latent heat", "latent")],
        ("E", "m", "L")),
    _eq("Pressure in a fluid", "pressure = height × density × gravitational field strength", "P = h × ρ × g", "P = h*rho*g",
        [("P", "P", "pressure", "pressure"), ("h", "h", "height of the column", "length"),
         ("rho", "ρ", "density", "density"), ("g", "g", "gravitational field strength", "gfield")]),
    _eq("Force on a conductor", "force = magnetic flux density × current × length", "F = B × I × l", "F = B*I*l",
        [("F", "F", "force", "force"), ("B", "B", "magnetic flux density", "fluxdensity"),
         ("I", "I", "current", "current"), ("l", "l", "length of conductor", "length")]),
]}


def plain_symbol(symbol):
    """E<sub>k</sub> -> Ek (for images and anywhere tags are not understood)."""
    return re.sub(r"</?sub>", "", symbol)
