from collections import Counter, defaultdict
from pathlib import Path
import re
import openpyxl

root = Path(r"C:\sites\sistema_clinica_edge\outputs\01a0d409-c7cb-7882-8379-5af976f25090")
source = Path(r"C:\Users\Usuário\Downloads\EMPRESAS BELEM (1) (1).xlsx")
requested = [line.strip() for line in (root / "requested_ids.txt").read_text(encoding="utf-8").splitlines() if line.strip()]
wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
ws = wb["EMPRESAS"]
records = [(r, *[ws.cell(r, c).value for c in range(1, 8)]) for r in range(2, 109)]
digits = lambda value: re.sub(r"\D", "", str(value or ""))
print("requested", len(requested), "source", len(records))
print("mismatches", [(i+1, requested[i], records[i][2]) for i in range(min(len(requested), len(records))) if digits(requested[i]) != digits(records[i][2])])
print("requested length 11/14", Counter(len(digits(v)) for v in requested))
print("source length 11/14", Counter(len(digits(v[2])) for v in records))
print("missing company", [(r, ident) for r, _, ident, company, *_ in records if not company])
print("missing email", [(r, ident) for r, _, ident, _, email, *_ in records if not email])
print("email with dash", [(r, ident, email) for r, _, ident, _, email, *_ in records if email and " - " in email])
print("source cc populated", [(r, ident, cc) for r, _, ident, _, _, cc, *_ in records if cc])
by_id = defaultdict(list)
for rec in records:
    by_id[digits(rec[2])].append(rec)
for ident, group in by_id.items():
    if len(group) > 1:
        print("duplicate", ident, [(r, company, email) for r, _, _, company, email, *_ in group])
