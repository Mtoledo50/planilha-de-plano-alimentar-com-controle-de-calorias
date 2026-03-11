from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import copy

wb = load_workbook('/home/claude/plano_alimentar.xlsx')

# ── Helpers ──────────────────────────────────────────
C = {
    "verde_esc":"1B5E20","verde_med":"2E7D32","verde_cla":"A5D6A7","verde_bg":"E8F5E9",
    "amarelo":"F57F17","amarelo_cla":"FFF9C4","amarelo_bg":"FFFDE7",
    "azul_med":"1565C0","azul_cla":"BBDEFB","azul_bg":"E3F2FD",
    "verm":"B71C1C","verm_cla":"FFCDD2","verm_bg":"FFEBEE",
    "laranja":"E65100","laranja_cla":"FFE0B2","laranja_bg":"FFF3E0",
    "teal":"004D40","teal_cla":"B2DFDB","teal_bg":"E0F2F1",
    "roxo":"4A148C","roxo_cla":"E1BEE7","roxo_bg":"F3E5F5",
    "cinza_esc":"212121","cinza_med":"424242","cinza":"757575",
    "cinza_cla":"EEEEEE","cinza_bg":"FAFAFA",
    "branco":"FFFFFF","input_am":"FFFDE7","input_bor":"F9A825","input_azul":"1A237E",
}

def F(h): return PatternFill("solid", start_color=h, fgColor=h)
def ft(sz=10, bold=False, color="212121", italic=False):
    return Font(name="Calibri", size=sz, bold=bold, color=color, italic=italic)
def al(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def s(style="thin", c="BDBDBD"): return Side(style=style, color=c)
def bdr(t="thin", c="BDBDBD"):
    sd = s(t, c); return Border(left=sd, right=sd, top=sd, bottom=sd)

BDR = bdr()
TBL = "'📋 Tabela de Alimentos'!$B$7:$J$300"

# ════════════════════════════════════════════════════════════
# STEP 1 — Add missing foods to 📋 Tabela de Alimentos
# ════════════════════════════════════════════════════════════
ws1 = wb['📋 Tabela de Alimentos']

# Collect existing food names
existing = set()
for row in ws1.iter_rows(min_row=8, max_row=ws1.max_row, min_col=2, max_col=2):
    v = row[0].value
    if v: existing.add(str(v).strip())

print("Existing foods:", len(existing))

# Category colors (same as build script)
cat_colors = {
    "🌾 Carboidratos / Cereais e Derivados": "F57F17",
    "🥩 Proteínas Animais — Carnes e Peixes": "B71C1C",
    "🫘 Proteínas Vegetais — Leguminosas":    "4E342E",
    "🥛 Laticínios e Ovos":                  "1565C0",
    "🥗 Saladas e Folhosos":                  "2E7D32",
    "🥦 Legumes e Verduras Cozidas":          "004D40",
    "🍎 Frutas":                              "880E4F",
    "💧 Bebidas":                             "0D47A1",
    "🥜 Oleaginosas, Sementes e Gorduras Boas":"E65100",
    "🍫 Suplementos e Funcionais":            "4A148C",
    "🍽️ Temperos e Condimentos":              "424242",
}

# New foods to add: (nome, categoria, kcal, prot, carb, gord, fibra, obs)
# Data from TACO/IBGE + TBCA/USP
NEW_FOODS = [
    # Requeijão light — TACO
    ("Requeijão light",   "🥛 Laticínios e Ovos",
     152, 7.5, 3.2, 12.0, 0.0,
     "TACO — aprox. 152kcal/100g. 15g = ~23 kcal"),
    # Peito de peru (presunto de peru fatiado) — TACO
    ("Peito de peru fatiado", "🥩 Proteínas Animais — Carnes e Peixes",
     109, 17.4, 1.9, 3.6, 0.0,
     "TACO — 1 fatia ≈ 15g = ~16 kcal. Baixo teor de gordura"),
    # Clara de ovo já existia, mas ovo inteiro precisamos verificar
    # Maçã com casca já está. Vamos adicionar "Maçã sem casca" e garantir "Maçã com casca"
    ("Maçã sem casca",    "🍎 Frutas",
     52,  0.2, 13.8, 0.2, 1.2,
     "TACO — 120g = ~62 kcal"),
    # Café com leite — Bebidas
    ("Café com leite desnatado", "💧 Bebidas",
     18,  1.4,  2.5, 0.1, 0.0,
     "Estimado — 200ml. Café + leite desnatado"),
    # Bebida vegetal de soja
    ("Leite de soja sem açúcar", "💧 Bebidas",
     33,  3.0,  2.7, 1.5, 0.3,
     "TBCA/USP"),
]

# Check which ones need to be added
to_add = [f for f in NEW_FOODS if f[0] not in existing]
print(f"Foods to add: {len(to_add)}: {[f[0] for f in to_add]}")

# Find the last row and where each category section ends
# We'll append new foods right after their category section
# Get category positions
cat_section_rows = {}  # cat_name -> last_row_of_section
current_cat = None
for row in ws1.iter_rows(min_row=7, max_row=ws1.max_row, min_col=2, max_col=3):
    b, c = row[0], row[1]
    if b.value and c.value is None:
        # section header
        current_cat = str(b.value).strip().lstrip()
        # clean up leading spaces
        for k in cat_colors:
            if k in current_cat or current_cat in k:
                current_cat = k
                break
    elif b.value and c.value and current_cat:
        cat_section_rows[current_cat] = b.row

print("Category last rows:", {k[:20]: v for k, v in cat_section_rows.items()})

# We'll insert each new food after its category's last row
# Since inserting rows shifts everything, we do it in reverse order of row number
# Group by category
from collections import defaultdict
by_cat = defaultdict(list)
for food in to_add:
    by_cat[food[1]].append(food)

# Sort categories by their last row descending so inserts don't shift other positions
sorted_cats = sorted(by_cat.keys(), key=lambda c: cat_section_rows.get(c, 999), reverse=True)

for cat in sorted_cats:
    foods_for_cat = by_cat[cat]
    insert_after = cat_section_rows.get(cat, ws1.max_row)
    bg_cat = cat_colors.get(cat, "757575")

    for j, (nome, categoria, kcal, prot, carb, gord, fibra, obs) in enumerate(foods_for_cat):
        insert_row = insert_after + 1 + j
        ws1.insert_rows(insert_row)
        ws1.row_dimensions[insert_row].height = 18
        alt_bg = "FAFAFA" if insert_row % 2 == 0 else "FFFFFF"

        # col B: name
        c = ws1.cell(row=insert_row, column=2, value=nome)
        c.font = ft(10); c.fill = F(alt_bg); c.border = BDR; c.alignment = al("left","center")
        # col C: category
        c = ws1.cell(row=insert_row, column=3, value=categoria)
        c.font = ft(8, color="FFFFFF"); c.fill = F(bg_cat); c.border = BDR; c.alignment = al("center","center",wrap=True)
        # cols D-H: nutritional values
        for col_i, val in [(4,kcal),(5,prot),(6,carb),(7,gord),(8,fibra)]:
            c = ws1.cell(row=insert_row, column=col_i, value=val)
            c.font = ft(10); c.fill = F(alt_bg); c.border = BDR; c.alignment = al("right","center"); c.number_format = "0.0"
        # col I: kcal/g formula
        c = ws1.cell(row=insert_row, column=9, value=f"=D{insert_row}/100")
        c.font = ft(10); c.fill = F(alt_bg); c.border = BDR; c.alignment = al("right","center"); c.number_format = "0.000"
        # col J: obs
        c = ws1.cell(row=insert_row, column=10, value=obs)
        c.font = ft(8, italic=True, color="757575"); c.fill = F(alt_bg); c.border = BDR; c.alignment = al("left","center")

        print(f"  Added '{nome}' at row {insert_row}")

# Get updated last food row
last_food_row = ws1.max_row
# Get all food names for dropdown (col B, skip section headers — check col C not None)
all_food_names = []
for row in ws1.iter_rows(min_row=8, max_row=last_food_row, min_col=2, max_col=3):
    name = row[0].value
    cat  = row[1].value
    if name and cat and str(cat).strip() not in ["", "Categoria"]:
        all_food_names.append(str(name).strip())

print(f"\nTotal foods in table after update: {len(all_food_names)}")

# ════════════════════════════════════════════════════════════
# STEP 2 — Create a hidden named list sheet for dropdown
# ════════════════════════════════════════════════════════════
# Add a hidden sheet "🔧 Listas" with all food names in column A
LIST_SHEET = "🔧 Listas"
if LIST_SHEET in wb.sheetnames:
    del wb[LIST_SHEET]

wsl = wb.create_sheet(LIST_SHEET)
wsl.sheet_state = 'hidden'
wsl.column_dimensions['A'].width = 40

wsl['A1'].value = "ALIMENTOS"
wsl['A1'].font = ft(9, True)

for i, name in enumerate(all_food_names):
    wsl.cell(row=i+2, column=1, value=name)

FOOD_LIST_RANGE = f"'{LIST_SHEET}'!$A$2:$A${len(all_food_names)+1}"
print(f"Food list range: {FOOD_LIST_RANGE}")

# ════════════════════════════════════════════════════════════
# STEP 3 — Update 🍽️ Plano do Dia — replace breakfast + all rows
# ════════════════════════════════════════════════════════════
ws2 = wb['🍽️ Plano do Dia']

# New breakfast items:
# café sem açúcar 200ml, pão francês 50g, requeijão light 15g,
# queijo mussarela 20g, peito de peru 1 fatia (15g),
# ovo inteiro cozido 2 unid (100g), maçã com casca 120g, creatina monohidratada 6g
NEW_CAFE = [
    ("Café sem açúcar",       200,  "200ml"),
    ("Pão francês",            50,  "50g = 1 unidade"),
    ("Requeijão light",        15,  "~1 col. sopa cheia"),
    ("Queijo mussarela",       20,  "1 fatia fina"),
    ("Peito de peru fatiado",  15,  "1 fatia"),
    ("Ovo inteiro cozido",    100,  "2 ovos ≈ 100g"),
    ("Maçã com casca",        120,  "1 maçã média"),
    ("Creatina monohidratada",  6,  "6g (2 doses)"),
]

# Current café da manhã occupies rows 9–13 (5 items after header at row 8)
# We need to replace those 5 rows with 8 new rows
# Header is at row 8, current items at rows 9-13

CAFE_HEADER_ROW = 8
OLD_CAFE_START  = 9
OLD_CAFE_END    = 13  # 5 rows
OLD_CAFE_COUNT  = OLD_CAFE_END - OLD_CAFE_START + 1
NEW_CAFE_COUNT  = len(NEW_CAFE)

# First, delete old café rows
for _ in range(OLD_CAFE_COUNT):
    ws2.delete_rows(OLD_CAFE_START)

# Insert new café rows
ws2.insert_rows(OLD_CAFE_START, amount=NEW_CAFE_COUNT)

# Style reference from remaining rows (get from what was row 14, now shifted)
# Breakfast style
CAFE_DARK  = C["amarelo"]
CAFE_LIGHT = C["amarelo_bg"]

TBL_REF = f"'📋 Tabela de Alimentos'!$B$7:$J$300"

for i, (alimento, porcao, medida) in enumerate(NEW_CAFE):
    row = OLD_CAFE_START + i
    ws2.row_dimensions[row].height = 20
    alt_bg = CAFE_LIGHT if i % 2 == 0 else "FFFFFF"

    # col B: seq number
    c = ws2.cell(row=row, column=2, value=i+1)
    c.font = ft(9, color=C["cinza"]); c.fill = F(alt_bg); c.border = BDR; c.alignment = al("center")

    # col C: Alimento (input, pre-filled — will get dropdown)
    c = ws2.cell(row=row, column=3, value=alimento)
    c.font = ft(10, bold=True, color=C["input_azul"]); c.fill = F(C["input_am"])
    c.border = bdr("thin", C["input_bor"]); c.alignment = al("left","center")

    # col D: Refeição
    c = ws2.cell(row=row, column=4, value="☀️ Café da Manhã")
    c.font = ft(9, color="FFFFFF"); c.fill = F(CAFE_DARK); c.border = BDR; c.alignment = al("center","center",wrap=True)

    # col E: Porção (g) — input
    c = ws2.cell(row=row, column=5, value=porcao)
    c.font = ft(10, bold=True, color=C["input_azul"]); c.fill = F(C["input_am"])
    c.border = bdr("thin", C["input_bor"]); c.alignment = al("right","center"); c.number_format = "0.0"

    # cols F-J: VLOOKUP formulas (kcal, prot, carb, gord, fibra)
    for col_f, offset in [(6,3),(7,4),(8,5),(9,6),(10,7)]:
        formula = f"=IFERROR(VLOOKUP(C{row},{TBL_REF},{offset},0)*E{row}/100,0)"
        c = ws2.cell(row=row, column=col_f, value=formula)
        c.font = ft(10); c.fill = F(alt_bg); c.border = BDR
        c.alignment = al("right","center"); c.number_format = "0.0"

    # col K: % meta calórica
    c = ws2.cell(row=row, column=11, value=f"=IFERROR(F{row}/$F$4,0)")
    c.font = ft(10); c.fill = F(alt_bg); c.border = BDR
    c.alignment = al("right","center"); c.number_format = "0.0%"

    # col L: medida referência (extra info)
    # Use column 12 for a note
    c = ws2.cell(row=row, column=12, value=medida)
    c.font = ft(8, italic=True, color=C["cinza"]); c.fill = F(alt_bg); c.border = BDR
    c.alignment = al("left","center")

# Widen column L for notes
ws2.column_dimensions['L'].width = 22

# Update the L header to "Medida / Referência"
ws2.cell(row=7, column=12, value="Medida Ref.")
ws2.cell(row=7, column=12).font = ft(9, True, "FFFFFF")
ws2.cell(row=7, column=12).fill = F(C["teal"])
ws2.cell(row=7, column=12).border = BDR
ws2.cell(row=7, column=12).alignment = al("center","center",wrap=True)
ws2.row_dimensions[7].height = 38

print(f"\nNew breakfast rows: {OLD_CAFE_START} to {OLD_CAFE_START + NEW_CAFE_COUNT - 1}")

# ════════════════════════════════════════════════════════════
# STEP 4 — Update sequence numbers for ALL food rows after breakfast
# ════════════════════════════════════════════════════════════
# After inserting rows, renumber all food rows (col B with numeric values)
seq = 1
for row in ws2.iter_rows(min_row=CAFE_HEADER_ROW+1, max_row=ws2.max_row, min_col=2, max_col=3):
    b, c = row[0], row[1]
    if isinstance(b.value, (int, float)) or (b.value and str(b.value).isdigit()):
        b.value = seq
        seq += 1

print(f"Renumbered {seq-1} food rows")

# ════════════════════════════════════════════════════════════
# STEP 5 — Add dropdown (Data Validation) to ALL alimento cells
# ════════════════════════════════════════════════════════════
# Find all food rows in ws2 (col B has a number, col C is food name)
food_cells_ws2 = []
for row in ws2.iter_rows(min_row=8, max_row=ws2.max_row, min_col=2, max_col=3):
    b, c = row[0], row[1]
    if b.value is not None and str(b.value).strip().isdigit():
        food_cells_ws2.append(c.coordinate)

print(f"Food cells to get dropdown: {len(food_cells_ws2)}")
print("Sample:", food_cells_ws2[:5])

# Create a single DataValidation for all food cells using the list sheet
# Excel dropdown: source must be a range reference or comma list
# We use the hidden sheet range
dv = DataValidation(
    type="list",
    formula1=FOOD_LIST_RANGE,
    allow_blank=True,
    showDropDown=False,   # False = show the arrow (counterintuitive but correct)
    showErrorMessage=True,
    errorTitle="Alimento inválido",
    error="Escolha um alimento da lista ou deixe em branco.",
    showInputMessage=True,
    promptTitle="🥗 Selecionar Alimento",
    prompt="Clique na seta ▼ para ver todos os alimentos disponíveis, ou digite para filtrar."
)
dv.sqref = " ".join(food_cells_ws2)

# Remove existing data validations first
ws2.data_validations.dataValidation = []
ws2.add_data_validation(dv)

print(f"Dropdown applied to {len(food_cells_ws2)} cells")

# ════════════════════════════════════════════════════════════
# STEP 6 — Update the "Calorias por Refeição" summary
# We need to find the new row ranges for each meal section
# ════════════════════════════════════════════════════════════
# Find meal section positions using column D (refeição label in food rows)
meal_food_rows = {}
for row in ws2.iter_rows(min_row=8, max_row=ws2.max_row, min_col=2, max_col=5):
    b, c_cell, d, e = row[0], row[1], row[2], row[3]
    # Food rows: col B is a number, col D has the meal name
    if b.value is not None and str(b.value).strip().isdigit() and d.value:
        meal_name = str(d.value).strip()
        if meal_name not in meal_food_rows:
            meal_food_rows[meal_name] = []
        meal_food_rows[meal_name].append(b.row)

print("\nMeal food rows (first/last):")
for m, rows in meal_food_rows.items():
    if rows:
        print(f"  {m}: rows {rows[0]}-{rows[-1]} ({len(rows)} items)")

# Find total row and summary rows
total_row = None
summary_start = None
for row in ws2.iter_rows(min_row=8, max_row=ws2.max_row, min_col=2, max_col=2):
    v = row[0].value
    if v and "TOTAL DO DIA" in str(v):
        total_row = row[0].row
    if v and "CALORIAS POR REFEIÇÃO" in str(v):
        summary_start = row[0].row + 2  # skip the header+column header rows

print(f"Total row: {total_row}, Summary start: {summary_start}")

# Update total row formulas
if total_row:
    first_food = min(min(rows) for rows in meal_food_rows.values() if rows)
    last_food  = max(max(rows) for rows in meal_food_rows.values() if rows)
    
    for col_t, num in [(6,"0.0"),(7,"0.0"),(8,"0.0"),(9,"0.0"),(10,"0.0")]:
        c = ws2.cell(row=total_row, column=col_t)
        c.value = f"=SUM(F{first_food}:F{last_food})" if col_t==6 else \
                  f"=SUM(G{first_food}:G{last_food})" if col_t==7 else \
                  f"=SUM(H{first_food}:H{last_food})" if col_t==8 else \
                  f"=SUM(I{first_food}:I{last_food})" if col_t==9 else \
                  f"=SUM(J{first_food}:J{last_food})"
        c.number_format = num

    # Update % meta
    ws2.cell(row=total_row, column=11).value = f"=IFERROR(F{total_row}/F4,0)"
    ws2.cell(row=total_row, column=11).number_format = "0.0%"

    # Update status row
    status_row = total_row + 1
    ws2.cell(row=status_row, column=6).value = f'=IF(F{total_row}<F4*0.9,"⚠️ Abaixo",IF(F{total_row}<=F4,"✅ Na meta",IF(F{total_row}<=F4*1.1,"🟡 Acima 10%","🔴 Excedido")))'
    ws2.cell(row=status_row, column=7).value = f'=IF(G{total_row}>=H5,"✅ OK",IF(G{total_row}>=H5*0.8,"🟡 Razoável","🔴 Baixo"))'
    ws2.cell(row=status_row, column=8).value = f'=IF(H{total_row}<=J5,"✅ OK",IF(H{total_row}<=J5*1.15,"🟡 Atenção","🔴 Excedido"))'

# Update summary section (per-meal calorie totals)
if summary_start and meal_food_rows:
    meal_order = ["☀️ Café da Manhã","🍌 Lanche da Manhã","🥃 Almoço","🍵 Lanche da Tarde","🌙 Jantar","🌛 Ceia"]
    meal_colors_local = {
        "☀️ Café da Manhã":   ("F57F17","FFF3E0"),
        "🍌 Lanche da Manhã": ("E65100","FFF3E0"),
        "🥃 Almoço":          ("2E7D32","E8F5E9"),
        "🍵 Lanche da Tarde": ("1565C0","E3F2FD"),
        "🌙 Jantar":          ("004D40","E0F2F1"),
        "🌛 Ceia":            ("4A148C","F3E5F5"),
    }
    r_sum = summary_start
    for j, meal in enumerate(meal_order):
        # find this row in ws2
        for sr in range(summary_start-1, ws2.max_row+1):
            c_val = ws2.cell(row=sr, column=2).value
            if c_val and meal.split()[0] in str(c_val) and meal.split()[-1] in str(c_val):
                r_sum = sr
                break
        
        rows_m = meal_food_rows.get(meal, [])
        if not rows_m:
            continue
        
        dark_m, light_m = meal_colors_local.get(meal, ("424242","FAFAFA"))
        
        # Update kcal, prot, carbs, % formulas
        row_range_f = f"F{rows_m[0]}:F{rows_m[-1]}"
        row_range_g = f"G{rows_m[0]}:G{rows_m[-1]}"
        row_range_h = f"H{rows_m[0]}:H{rows_m[-1]}"

        c = ws2.cell(row=r_sum, column=4)
        c.value = f"=SUM({row_range_f})"
        c.number_format = "0.0"

        c = ws2.cell(row=r_sum, column=6)
        c.value = f"=SUM({row_range_g})"
        c.number_format = "0.0"

        c = ws2.cell(row=r_sum, column=8)
        c.value = f"=SUM({row_range_h})"
        c.number_format = "0.0"

        # % of daily total
        if total_row:
            c = ws2.cell(row=r_sum, column=10)
            c.value = f"=IFERROR(D{r_sum}/F{total_row},0)"
            c.number_format = "0.0%"

print("\nSummary section updated")

# ════════════════════════════════════════════════════════════
# STEP 7 — Update Dashboard formulas with new row references
# ════════════════════════════════════════════════════════════
ws3 = wb['📊 Dashboard de Macros']
if total_row:
    # Update KPI cards in dashboard (rows 5 = value row)
    new_formulas = {
        2: f"='🍽️ Plano do Dia'!F{total_row}",
        3: "='🍽️ Plano do Dia'!F4",
        4: f"='🍽️ Plano do Dia'!G{total_row}",
        5: f"='🍽️ Plano do Dia'!H{total_row}",
        6: f"='🍽️ Plano do Dia'!I{total_row}",
        7: f"='🍽️ Plano do Dia'!J{total_row}",
    }
    for col, formula in new_formulas.items():
        c = ws3.cell(row=5, column=col)
        c.value = formula

    # Update macro distribution rows (rows 9-11, col 3)
    for r in range(9, 12):
        c3 = ws3.cell(row=r, column=3)
        if c3.value and "Plano" in str(c3.value):
            macro_map = {
                9: f"='🍽️ Plano do Dia'!G{total_row}",
                10: f"='🍽️ Plano do Dia'!H{total_row}",
                11: f"='🍽️ Plano do Dia'!I{total_row}",
            }
            c3.value = macro_map.get(r, c3.value)

    # Update indicators
    ff = first_food if total_row else 9
    lf = last_food if total_row else 40
    for r in range(1, ws3.max_row+1):
        c3 = ws3.cell(row=r, column=3)
        v = str(c3.value or "")
        if "Plano do Dia" in v and "tot_row" in v.lower():
            c3.value = v.replace("tot_row", str(total_row))

print("Dashboard updated")

# ════════════════════════════════════════════════════════════
# STEP 8 — Freeze and finalize
# ════════════════════════════════════════════════════════════
wb.active = ws2
ws2.freeze_panes = "B8"

wb.save('/home/claude/plano_alimentar_v2.xlsx')
print("\n✅ Saved: plano_alimentar_v2.xlsx")
print(f"   - {len(all_food_names)} alimentos na tabela")
print(f"   - {len(NEW_CAFE)} itens no café da manhã")
print(f"   - Dropdown em {len(food_cells_ws2)} células")
