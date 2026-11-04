"""Genera data/*.csv y data/data.json desde 'Base de datos.xlsx' (sin nombres de personas)."""
import json, re, pandas as pd
from pathlib import Path
here = Path(__file__).parent
src = here.parent / "Base de datos.xlsx"
x = pd.read_excel(src, sheet_name=None)
for df in x.values():  # espacios sobrantes duplicaban instituciones (" Romeral" vs "Romeral")
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].str.strip().str.replace(r"\s+", " ", regex=True)

b = x["Base de datos"].copy()
# unifica variantes de mayúsculas ("Bienestar universitario" / "Bienestar Universitario")
b["Actividad"] = b["Actividad"].str.lower().map(b.groupby(b["Actividad"].str.lower())["Actividad"].agg(lambda s: s.mode()[0]))
b["Fecha"] = b["Fecha"].dt.strftime("%Y-%m-%d")
b = b.rename(columns={"Institucion": "Institución", "Convenio": "Financiador"})
b["Cantidad"] = b["Cantidad"].fillna(0)
b["Valor"] = b["Valor"].fillna(0).round(0).astype(int)

t = x["Técnicos"].copy()
t.columns = [c.strip() for c in t.columns]
t = t.drop(columns=["Nombres y Apellidos"])  # dato personal: no se publica
t["Programa"] = (t["Programa"].str.strip()
                 .str.replace(r"^T\.P\.", "Técnico Profesional", regex=True)
                 .str.replace("Matenimiento", "Mantenimiento").str.replace("Cofiguración", "Configuración"))
t["Año Graduación"] = t["Año Graduación"].astype("Int64")

e = x["Empresarismo"].copy().drop(columns=["Nombre"])
e["Tipo de proyecto"] = e["Tipo de proyecto"].replace({"Apícultura": "Apícola", "Turistico": "Turístico"})

d = here / "data"
for name, df in [("inversion", b), ("tecnicos", t), ("empresarismo", e)]:
    df.to_csv(d / f"{name}.csv", index=False, encoding="utf-8-sig")  # BOM: Excel lee tildes bien
    print(name, len(df))

out = {k: json.loads(v.to_json(orient="records", force_ascii=False, date_format="iso"))
       for k, v in [("inv", b), ("tec", t), ("emp", e)]}

# Meta 2026: filas de componente (nombre en MAYÚSCULAS) con "VR TOTAL PROPUESTA" (col F) en la hoja SGTO
sg = pd.read_excel(here.parent / "SGTO Fund Bancolombia 2026.xlsx", header=None)
NOMBRES = {"ESTRATEGIAS DE PERMANENCIA": ("Estrategias de permanencia", "Estrategia de permanencia"),
           "PRUEBAS SABER T Y T": ("Pruebas Saber T y T", None)}  # (nombre a mostrar, Proceso en la base)
meta = []
for _, r in sg.iterrows():
    n, v = r[0], pd.to_numeric(r[5], errors="coerce")
    if isinstance(n, str) and n.strip() == n.strip().upper() and pd.notna(v):
        n = n.strip()
        nom, proc = NOMBRES.get(n, (n.capitalize(), n.capitalize()))
        meta.append({"componente": nom, "proceso": proc, "meta": int(round(v))})
assert meta, "No se encontró la meta en el SGTO: revisar formato de la hoja"
out["meta"] = meta
print("meta", sum(m["meta"] for m in meta), [m["componente"] for m in meta])
(d / "data.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
