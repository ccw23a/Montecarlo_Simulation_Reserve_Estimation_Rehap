#!/usr/bin/env python3
"""Monte Carlo port of Excel 'Reserve Calculation' sheet.
Reads key parameters from the XLSX using stdlib (no external deps), runs simulations, writes CSV.
"""
import zipfile
import xml.etree.ElementTree as ET
import math
import csv
from datetime import datetime

XLSX_PATH = 'Excel_Base/Reserve Calculation_Monte Carlo_2026.xlsx'

# --- Helpers to read cell values from xlsx (stdlib only) ---

def load_workbook_xml(xlsx_path):
    z = zipfile.ZipFile(xlsx_path)
    return z

def build_shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return {}
    data = z.read('xl/sharedStrings.xml')
    root = ET.fromstring(data)
    s = {}
    for i, si in enumerate(root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')):
        t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
        s[i] = t.text if t is not None else ''
    return s

def map_sheets(z):
    wb = z.read('xl/workbook.xml')
    root = ET.fromstring(wb)
    ns = {'ns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheets = root.find('ns:sheets', ns)
    name_to_rid = {}
    for sh in sheets.findall('ns:sheet', ns):
        name = sh.attrib.get('name')
        rid = sh.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or sh.attrib.get('r:id')
        name_to_rid[name] = rid
    # rels
    rels = z.read('xl/_rels/workbook.xml.rels')
    rroot = ET.fromstring(rels)
    rid_to_target = {}
    for rel in rroot.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid_to_target[rel.attrib.get('Id')] = rel.attrib.get('Target')
    # build final map sheet -> path
    m = {}
    for n, rid in name_to_rid.items():
        target = rid_to_target.get(rid)
        if target:
            m[n] = 'xl/' + target
    return m

def load_sheet_cells(z, sheet_path, shared_strings):
    data = z.read(sheet_path)
    root = ET.fromstring(data)
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    cells = {}
    for c in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        coord = c.attrib.get('r')
        t = c.attrib.get('t')
        f = c.find(ns + 'f')
        v = c.find(ns + 'v')
        if f is not None:
            cells[coord] = {'type':'formula', 'text': f.text}
        elif v is not None:
            val = v.text
            if t == 's':
                val = shared_strings.get(int(val), val)
            else:
                # try convert to float/int
                try:
                    if '.' in val or 'E' in val or 'e' in val:
                        val = float(val)
                    else:
                        val = int(val)
                except Exception:
                    pass
            cells[coord] = {'type':'value', 'value': val}
    return cells

# --- Numeric utilities ---

def norm_ppf(p):
    # Approximation by Peter John Acklam (implemented as commonly used)
    # Valid for 0<p<1
    if p <= 0.0 or p >= 1.0:
        raise ValueError('p must be in (0,1)')
    # Coefficients
    a = [ -3.969683028665376e+01,  2.209460984245205e+02,
          -2.759285104469687e+02,  1.383577518672690e+02,
          -3.066479806614716e+01,  2.506628277459239e+00 ]
    b = [ -5.447609879822406e+01,  1.615858368580409e+02,
          -1.556989798598866e+02,  6.680131188771972e+01,
          -1.328068155288572e+01 ]
    c = [ -7.784894002430293e-03, -3.223964580411365e-01,
          -2.400758277161838e+00, -2.549732539343734e+00,
           4.374664141464968e+00,  2.938163982698783e+00 ]
    d = [ 7.784695709041462e-03,  3.224671290700398e-01,
          2.445134137142996e+00,  3.754408661907416e+00 ]
    plow = 0.02425
    phigh = 1 - plow
    if p < plow:
        q = math.sqrt(-2*math.log(p))
        x = (((((c[0]*q + c[1])*q + c[2])*q + c[3])*q + c[4])*q + c[5]) / (
             (((d[0]*q + d[1])*q + d[2])*q + d[3])*q + 1)
    elif p > phigh:
        q = math.sqrt(-2*math.log(1-p))
        x = -(((((c[0]*q + c[1])*q + c[2])*q + c[3])*q + c[4])*q + c[5]) / (
              (((d[0]*q + d[1])*q + d[2])*q + d[3])*q + 1)
    else:
        q = p - 0.5
        r = q*q
        x = (((((a[0]*r + a[1])*r + a[2])*r + a[3])*r + a[4])*r + a[5])*q / (
             (((((b[0]*r + b[1])*r + b[2])*r + b[3])*r + b[4])*r + 1))
    return x

def clamp_norm_inv(u, mean, sd, mn, mx):
    z = norm_ppf(u)
    val = mean + sd * z
    if val < mn:
        return mn
    if val > mx:
        return mx
    return val

def triangular_inv(u, a, c, b):
    # a=min, c=mode, b=max
    if b == a:
        return a
    fc = (c - a) / (b - a)
    if u < fc:
        return a + math.sqrt(u * (c - a) * (b - a))
    else:
        return b - math.sqrt((1 - u) * (b - a) * (b - c))

# --- Main: read params, run sims ---

def get_cell_value(cells, coord):
    ent = cells.get(coord)
    if not ent:
        return None
    if ent['type'] == 'value':
        return ent['value']
    return None


def main(n_sims=2000, out_csv='outputs/montecarlo_results.csv'):
    z = load_workbook_xml(XLSX_PATH)
    shared = build_shared_strings(z)
    sheet_map = map_sheets(z)
    sheet_path = sheet_map.get('Reserve Calculation')
    if not sheet_path:
        raise RuntimeError('Reserve Calculation sheet not found')
    cells = load_sheet_cells(z, sheet_path, shared)

    # Read parameter cells (as observed from the workbook formulas)
    B4 = float(get_cell_value(cells,'B4'))
    B5 = float(get_cell_value(cells,'B5'))
    C4 = float(get_cell_value(cells,'C4'))
    C5 = float(get_cell_value(cells,'C5'))
    C6 = float(get_cell_value(cells,'C6'))
    D4 = float(get_cell_value(cells,'D4'))
    D5 = float(get_cell_value(cells,'D5'))
    D6 = float(get_cell_value(cells,'D6'))
    F4 = float(get_cell_value(cells,'F4'))
    F5 = float(get_cell_value(cells,'F5'))

    results = []
    for i in range(n_sims):
        u = (i+1) / (n_sims+1)  # deterministic quasi-random like Excel RAND per row; using uniform sequence to be deterministic
        # Excel uses RAND() per row; to mimic variety use random.random() instead of sequence if desired.
        # Use deterministic to match reproducible outputs; user can switch to random.random()
        Vb = clamp_norm_inv(u, mean=(B5+B4)/2.0, sd=(B5-B4)/6.0, mn=B4, mx=B5)
        d_val = triangular_inv(u, a=C4, c=C6, b=C5)
        e_val = triangular_inv(u, a=D4, c=D6, b=D5)
        f_val = 1.0 - e_val
        g_val = 1.2
        h_val = clamp_norm_inv(u, mean=(F5+F4)/2.0, sd=(F5-F4)/6.0, mn=F4, mx=F5)
        reserve = 7758.0 * Vb * d_val * f_val * h_val / g_val
        results.append((Vb, d_val, e_val, f_val, g_val, h_val, reserve))

    # write CSV
    with open(out_csv, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Vb','param_d','Sw','So','G','H','Reserve'])
        for row in results:
            w.writerow(row)

    print(f'Wrote {len(results)} rows to {out_csv}')

if __name__ == '__main__':
    main()
