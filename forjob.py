import zipfile
import json
import csv
from pathlib import Path

# pbix. на одном уровне со скриптом (общая папка)

# === Настройки ===
pbix_path = "ДБ Конкуренты v1.4.2.pbix"             # путь к PBIX-файлу - Единственное поле, которое нужно менять.
output_csv = "visual_fields_" + pbix_path + ".csv"      # куда сохранить результат

# === Временная распаковка ===
tmp_dir = Path("pbix_tmp")
tmp_dir.mkdir(exist_ok=True)

with zipfile.ZipFile(pbix_path, 'r') as z:
    z.extractall(tmp_dir)

layout_path = tmp_dir / "Report" / "Layout"

if not layout_path.exists():
    raise FileNotFoundError("Файл Layout не найден. Убедись, что PBIX не защищён и структура стандартная.")

# === Универсальное чтение JSON с автоопределением кодировки ===
def load_json_safely(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-16", "utf-16-le", "utf-16-be", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            return json.loads(text)
        except Exception:
            continue
    raise ValueError(f"Не удалось определить кодировку файла: {path}")

layout_json = load_json_safely(layout_path)

# === Извлечение визуалов ===
results = []

for page in layout_json.get("sections", []):
    page_name = page.get("displayName", "Без названия")
    visuals = page.get("visualContainers", [])
    for vis in visuals:
        config_str = vis.get("config", "")
        try:
            config_json = json.loads(config_str)
        except Exception:
            continue

        visual_type = config_json.get("singleVisual", {}).get("visualType", "Unknown")

        # --- Поиск полей ---
        fields = set()

        # 1. prototypeQuery → Select
        selects = config_json.get("singleVisual", {}).get("prototypeQuery", {}).get("Select", [])
        for s in selects:
            if "Name" in s:
                fields.add(s["Name"])

        # 2. dataRoles → members → field.Column.Property
        roles = config_json.get("singleVisual", {}).get("dataRoles", [])
        for role in roles:
            for m in role.get("members", []):
                col = m.get("field", {}).get("Column", {})
                tbl = col.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
                prop = col.get("Property", "")
                if tbl and prop:
                    fields.add(f"{tbl}[{prop}]")

        # 3. Если визуал slicer, KPI, карта и т.д. — часто поля только в prototypeQuery
        if not fields and "prototypeQuery" in config_json.get("singleVisual", {}):
            pq = config_json["singleVisual"]["prototypeQuery"]
            for key, val in pq.items():
                if isinstance(val, dict) and "Name" in val:
                    fields.add(val["Name"])

        for f in filter(None, fields):
            results.append({
                "Page": page_name,
                "VisualType": visual_type,
                "Field": f
            })

# === Сохранение результатов ===
with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=["Page", "VisualType", "Field"])
    writer.writeheader()
    writer.writerows(results)

print(f"✅ Найдено записей: {len(results)}")
print(f"📄 Результат сохранён в: {output_csv}")
