import openpyxl
import re
from openpyxl.utils.cell import rows_from_range

__all__ = ['parsing_formula']


class Constants_dict():
    """
    klasa ma przechowywac stale w dict CONSTnr: wartosc
    """

    def __init__(self) -> None:
        self.my_dict = dict()
        self.constant_id = 0

    def add_constant(self, val):
        self.constant_id += 1
        self.my_dict['CONST' + str(self.constant_id)] = val
        return 'CONST' + str(self.constant_id)


def _parsed_inseide_sum(s):
    """
    Pomocnicza funkcja wywowała w _replace_sum
    rozbija A1:A44 na sume komorek A1+A2 ...
    """
    new_f = []
    for rng in s.split(','):
        cell_in_rng = list(rows_from_range(rng))
        if len(cell_in_rng) == 1:
            new_f.extend(cell_in_rng[0])
        else:
            new_f.extend([x[0] for x in cell_in_rng])
    return(new_f)


def _replace_sum(s):
    """
    pasruje SUM(...)
    zwaraca (A1 .... + A33)
    """
    pattern = r"SUM\((?P<in_sum>[A-Z0-9:,]+)\)"
    match_iter = re.finditer(pattern, s)
    for match in match_iter:
        to_be_replaced = match.group(0)
        in_sum = match.group('in_sum')
        new_f = _parsed_inseide_sum(in_sum)
        replace_with = '(' + '+'.join(new_f) + ')'
        s = s.replace(to_be_replaced, replace_with)
    return s


def _shift_right_multiplier_of_cell_to_left(f):
    """
    Dla kazdej komorki przenosi mnozik z prawkej na lewa
    np z A3*5 -> 5*A3
    """
    regex = r"(?P<CELL>[A-Z]+\d*)(?:\*(?P<Po>[+-]?(?:\d*\.\d+|\d+\.\d*|\d+)))"
    subst = "\\g<Po>*\\g<CELL>"
    result = re.sub(regex, subst, f, 0, re.MULTILINE)
    return result


def _clean_operators(s):
    s = s.replace("+-", "-").replace('-+',
                                     '-').replace('--', '+').replace('++', '+').replace(' ', '')
    return s


def _shift_right_multiplier_to_left(f):
    """
    Przenoszi mnozik z prawje strony nawiasu na lewa.
    Pierwotnie ta funckja czyscila zapis sum(A1:A2)*5
    Ale wykorzytuje ja do wszystkch nawiasow
    """
    # mnoznik z prawej na lewa
    regex = r"(?P<Przed>(?:[\+\-\(])|(?:\d*\.\d+|\d+\.\d*|\d+\*)|(?:[\+\-\(]\d*\.\d+|\d+\.\d*|\d+\*))?(?P<W_NAWIASIE>\([^()]*\))(?:\*(?P<Po>[+-]?(?:\d*\.\d+|\d+\.\d*|\d+)))?"
    subst = r"\g<Przed>\g<Po>\g<W_NAWIASIE>"
    result = re.sub(regex, subst, f, 0, re.MULTILINE)

    # np 2(, dodanie * pomiedzy
    regex = r"(?P<Liczba>\d*\.\d+|\d+\.\d*|\d+)(?P<nawias>\()"
    subst = "\\g<Liczba>*\\g<nawias>"
    result = re.sub(regex, subst, result, 0, re.MULTILINE)

    result = _clean_operators(result)
    result = _eval_for_two_multipliers(result)

    return result


def _eval_for_two_multipliers(f):
    """
    Gdy mamy dwa operatory np +2*4* -> +8*
    """
    def _m_eval(m):
        """
        to jest callback w funckji sub
        m to obiekt match
        """
        prefix = ''
        mg = m.group('Mnoznik')
        if mg[0] == '=':
            mg = mg[1:]
            prefix = '='
        if mg[0] == '(':
            mg = mg[1:]
            prefix = '('
        e = eval(mg)
        if e > 0:
            ev = prefix + '+' + str(e)
        else:
            ev = prefix + str(e)
        return ev

    # eval dwoch mnoznikow
    regex = r"(?P<Mnoznik>(?:[+-]?[^A-Z](?:\d*\.\d+|\d+\.\d*|\d+))\*[+-]?(?:\d*\.\d+|\d+\.\d*|\d+))"
    result = re.sub(regex, _m_eval, f, 0, re.MULTILINE)
    result = _clean_operators(result)
    return result


def _remove_parentheeses(s, start, end):
    """
    To jest odpalane w rekursi i robi duzo
    Usuwa nawiasy
    """
    s = _shift_right_multiplier_to_left(s)
    x = s[start:end]

    next_char = ''
    if end < len(s):
        next_char = s[end]

    # pattern +(....)+-)  >-just remove parentheeses
    if s[start-1] in {'+', '(', '='} and next_char in {'+', '-', ')', ''}:
        new_x = x.replace('(', '').replace(')', '')
        s = s[:start] + new_x + s[end:]
        return s

    # pattern -(....)+-)  >-,multiplie by -1
    if s[start-1] == '-' and next_char in {'+', '-', ')', ''}:
        new_x = x.replace('(', '').replace(')', '')
        if new_x[0] != '-':
            new_x = '+' + new_x
        new_x = new_x.replace('+', '__+__').replace('-',
                                                    '+').replace('__+__', '-')
        # new_x = '+' + new_x
        # new_x = _clean_operators(new_x)
        # s = s.replace(x, new_x)
        s = s[:start-1] + new_x + s[end:]
        return s

# /(?P<Liczba_pzred>\d+\*)?(?P<operator>-?)(?P<cell>[A-Z].+?)(?P<Liczba_po>\*[-+]\d+)?/gm
# \g<operator>2*\g<Liczba_pzred>\g<cell>\g<Liczba_po>
    # patter N * (....) -> (N*. + N*. ...)
    if next_char in {'+', '-', ')', ''}:
        on_left = s[:start]
        matches = re.search(
            r"(?P<m_before>[+-]?(?:\d*\.\d+|\d+\.\d*|\d+))\*$", on_left)
        if matches:
            multiplier = matches.group('m_before')
            new_x = x.replace('(', '').replace(')', '')
            regex = r"(?P<Liczba_pzred>\d+\*)?(?P<operator>-?)(?P<cell>[A-Z]{1,3}\d+)(?P<Liczba_po>\*[-+]\d+)?"
            subst = "\\g<operator>" + \
                str(multiplier) + "*\\g<Liczba_pzred>\\g<cell>\\g<Liczba_po>"
            result = re.sub(regex, subst, new_x, 0, re.MULTILINE)
            result = _eval_for_two_multipliers(result)
            result = _clean_operators(result)
            s = s.replace(multiplier + '*' + x, result)
            return s

    return s


def _read_parentheeses(f, iter=1):
    if iter > 200:
        return f

    if "(" not in f:
        return f

    m = re.search(r"\([^()]*\)", f)
    f = _remove_parentheeses(f, m.start(), m.end())

    if "(" in f or iter <= 4:
        f = _read_parentheeses(f, iter+1)
    return f


def _compress_formula(f):
    f = _eval_for_two_multipliers(f)
    regex = r"(?P<Znak>[+-])?(?P<Liczba>(?:\d*\.\d+|\d+\.\d*|\d+))?\*?(?P<CELL>[A-Z]+\d*)"
    formula = {}
    matches = re.finditer(regex, f, re.MULTILINE)
    for match in matches:
        cell = match.group('CELL')
        multipl = match.group('Liczba') or 1
        znak = match.group('Znak') or '+'
        formula[cell] = formula.get(cell, 0) + eval(znak+str(multipl))
    result = '=' + '+'.join([str(v) + '*' + k for k, v in formula.items()])
    result = _clean_operators(result)

    return result


def _compress_constants(f):
    _const = 0
    regex = r"[+-=](?:\d*\.\d+|\d+\.\d*|\d+)\*(?P<Liczba>(?:\d*\.\d+|\d+\.\d*|\d+))"
    matches = re.finditer(regex, f, re.MULTILINE)
    for m in matches:
        _const += eval(m.group())
    new_const = '+' + str(_const) if _const > 0 else str(_const)

    regex = r"[+-=](?:\d*\.\d+|\d+\.\d*|\d+)\*(?:(?:\d*\.\d+|\d+\.\d*|\d+))"
    result = re.sub(regex, "", f, 0, re.MULTILINE)

    if _const != 0:
        result = result + new_const
    return result


def _constants(f, d):
    regex = r"(?:[^\*][-]|[+])((?:\d*\.\d+|\d+\.\d*|\d+))[+-]|(?:[^\*][-]|[+])((?:\d*\.\d+|\d+\.\d*|\d+))$|^[-]?((?:\d*\.\d+|\d+\.\d*|\d+))|\([-]?((?:\d*\.\d+|\d+\.\d*|\d+))[+-]|(?:[^\*]-|[+])((?:\d*\.\d+|\d+\.\d*|\d+))\)|=[-]?((?:\d*\.\d+|\d+\.\d*|\d+))[+-]"
    m = re.search(regex, f, re.MULTILINE)
    while True:
        m = re.search(regex, f, re.MULTILINE)
        if not m:
            break
        for groupNum in range(0, len(m.groups())):
            groupNum = groupNum + 1
            if m.group(groupNum) is not None:
                x = d.add_constant(m.group(groupNum))
                f = f[:m.start(groupNum)] + x + f[m.end(groupNum):]
    return f


def parsing_formula(f):
    const_dict = Constants_dict()

    m = re.findall('(?:\d*\.\d+|\d+\.\d*|\d+)%', f)
    if m:
        for l in m:
            to_replace = l
            num = round(float(l[:-1]) / 100, 4)
            f = f.replace(to_replace, str(num))

    f1 = _replace_sum(f)
    f2 = _clean_operators(f1)
    _f2 = _constants(f2, const_dict)
    f3 = _shift_right_multiplier_of_cell_to_left(_f2)
    f4 = _shift_right_multiplier_to_left(f3)
    f5 = _eval_for_two_multipliers(f4)
    f6 = _read_parentheeses(f5)
    f7 = _compress_formula(f6)
    for k, v in const_dict.my_dict.items():
        f7 = f7.replace(k, v)
    f8 = _compress_constants(f7)

    if not (f8[1] in {'+', '-'}):
        f8 = "=+" + f8[1:]
    pattern = r"(?P<CellProduct>[+-](?:\d*\.\d+|\d+\.\d*|\d+)\*(?:[A-Z]+\d*))|(?P<Const>[+-](?:\d*\.\d+|\d+\.\d*|\d+)$)"
    frm = []
    fnd = re.findall(pattern, f8)
    for x in fnd:
        cell = x[0]+x[1]
        if '*' not in cell:
            cell = cell[0] + "1*" + cell[1:]
        frm.append(cell)

    return frm
