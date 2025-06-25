from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json

def load_json(json_path):
    with open(json_path, "r") as f:
        return json.load(f)

def is_number(value):
    try:
        float(value)
        return True
    except Exception:
        return False

def format_currency(value, decimals=2):
    try:
        value = round(float(value), decimals)
        if decimals == 0:
            return "${:,.0f}".format(value)
        else:
            return "${:,.2f}".format(value)
    except Exception:
        return str(value)

def calculate_broker_fee(base, broker_fee_pct, commission_pct, commission_without_gst):
    try:
        commission_without_gst_val = float(commission_without_gst)
    except (ValueError, TypeError):
        commission_without_gst_val = 0

    # Calculate shortfall in commission
    commission_shortfall_pct = max(commission_pct - commission_without_gst_val / base * 100 if base else 0, 0)

    # Add the shortfall to the broker fee
    effective_broker_fee_pct = broker_fee_pct + commission_shortfall_pct

    return round(base * (effective_broker_fee_pct / 100.0), 2)


def enrich_insurer_quotes(quotes_dict, broker_fee_pct, commission_pct, associate_split=0):
    enriched = {}
    for insurer, quote in quotes_dict.items():
        base = float(quote.get("base", 0) or 0)
        total = float(quote.get("total", 0) or 0)
        commission_without_gst = quote.get("commission_without_gst", 0) or 0
        try:
            commission_without_gst_val = float(commission_without_gst)
        except Exception:
            commission_without_gst_val = 0

        broker_fee = calculate_broker_fee(base, broker_fee_pct, commission_pct, commission_without_gst_val)
        broker_gst = round(broker_fee * 0.1, 2)
        remuneration = round(commission_without_gst_val + broker_fee, 2)
        final_total = round(total + broker_fee + broker_gst, 2)

        sm_remuneration = round(remuneration * (associate_split / 100), 2)
        broker_remuneration = round(remuneration - sm_remuneration, 2)

        enriched_quote = dict(quote)
        enriched_quote["broker_fee"] = broker_fee
        enriched_quote["broker_gst"] = broker_gst
        enriched_quote["remuneration"] = remuneration
        enriched_quote["sm_remuneration"] = sm_remuneration
        enriched_quote["broker_remuneration"] = broker_remuneration
        enriched_quote["final_total"] = final_total
        enriched_quote["_final_total_numeric"] = final_total
        enriched_quote["insurer"] = insurer

        enriched[insurer] = enriched_quote
    return enriched

def find_recommended(enriched_quotes):
    min_total = float("inf")
    best = None
    for insurer, quote in enriched_quotes.items():
        total_val = quote.get("_final_total_numeric", float("inf"))
        if total_val and total_val < min_total:
            min_total = total_val
            best = quote
    if best:
        result = dict(best)
        result.pop("_final_total_numeric", None)
        return result
    return {}

def flatten_data_for_replace(data, broker_fee_pct, commission_pct,strata_manager):
    flat = {}
    for k, v in data.get("general_info", {}).items():
        flat[k] = v
    associate_split = data.get("associate_split", 0)
    enriched_quotes = enrich_insurer_quotes(data.get("Quotes", {}), broker_fee_pct, commission_pct, associate_split)
    for insurer, insurer_data in enriched_quotes.items():
        for field, value in insurer_data.items():
            if not field.startswith("_"):
                if is_number(value) and field not in ["uwgst", "uw", "uwgst_fee"]:
                    flat[f"{insurer}.{field}"] = format_currency(value)
                else:
                    flat[f"{insurer}.{field}"] = value
    recommended = find_recommended(enriched_quotes)
    for field, value in recommended.items():
        if is_number(value):
            flat[f"recommended.{field}"] = format_currency(value)
        else:
            flat[f"recommended.{field}"] = value
    flat["broker_fee_pct"] = f"{broker_fee_pct}%"
    flat["commission_pct"] = f"{commission_pct}%"
    flat["strata_manager"] = strata_manager
    return flat

def set_cell_background(cell, rgb_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), rgb_color)
    tcPr.append(shd)

def set_cell_bottom_border(cell, color="357ABD", size="12"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:color'), color)
    borders.append(bottom)

def ensure_landscape_section(doc):
    section = doc.sections[-1]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = 1  # Landscape
    section.page_width = new_width
    section.page_height = new_height

def insert_market_summary_table(doc, quotes, recommended_insurer):
    placeholder = "{{market_summary_table}}"
    insurers = [
        "Axis", "CHU", "Flex", "Hutch", "IIS", "Longitude", "QUS", "SCI", "SUU"
    ]
    underwriters = {
        "Axis": "XL Insurance Company",
        "CHU": "QBE Insurance (Australia) Limited",
        "Flex": "QBE Insurance (Australia) Limited",
        "Hutch": "Certain Underwriters at Lloyds of London",
        "IIS": "Certain Underwriters at Lloyds of London",
        "Longitude": "Chubb Insurance Australia Limited",
        "QUS": "Certain Underwriters at Lloyds of London",
        "SCI": "Allianz Australia Insurance Limited",
        "SUU": "CGU Insurance Limited"
    }
    col_widths = [Inches(2.5), Inches(2.0), Inches(2.5)]

    for i, p in enumerate(doc.paragraphs):
        if placeholder in p.text:
            parent = p._element.getparent()
            idx = parent.index(p._element)
            parent.remove(p._element)

            tbl = doc.add_table(rows=1 + len(insurers), cols=3)
            tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            tbl.autofit = True

            headers = ["Insurer / Underwriter", "Premium Payable", "Comment"]
            for col in range(3):
                cell = tbl.cell(0, col)
                
                cell.text = headers[col]
                set_cell_background(cell, "FFFFFF")
                set_cell_bottom_border(cell)
                for p in cell.paragraphs:
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    for r in p.runs:
                        r.font.bold = True
                        r.font.size = Pt(9)
                        r.font.name = "Futura Bk BT (Body)"
                        r.font.color.rgb = RGBColor(0, 0, 0)

            for row_idx, insurer in enumerate(insurers, 1):
                color = "e9edf7" if row_idx % 2 == 1 else "FFFFFF"
                data = quotes.get(insurer)
                premium = None
                comment = "Insurer did not respond in time"
                if data:
                    enriched = enrich_insurer_quotes({insurer: data}, 20, 20)
                    premium = enriched[insurer].get("final_total")
                    comment = "Recommended" if insurer == recommended_insurer else ""
                insurer_label = f"{insurer} – Underwritten by {underwriters.get(insurer, 'Unknown')}"
                cells = [insurer_label, format_currency(premium) if premium else "", comment]

                for col_idx, value in enumerate(cells):
                    cell = tbl.cell(row_idx, col_idx)
                    
                    cell.text = value if value else ""
                    set_cell_background(cell, color)
                    for p in cell.paragraphs:
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                        for r in p.runs:
                            r.font.size = Pt(9)
                            r.font.name = "Futura Bk BT (Body)"

            parent.insert(idx, tbl._element)
            break

# Modified insert_comparison_table with landscape page width support

def insert_comparison_table(doc, quotes):
    ensure_landscape_section(doc)
    placeholder = "{{comparison_table}}"
    insurer_list = list(quotes.keys())
    first_features = next(iter(quotes.values())).get("features", {})
    feature_keys = list(first_features.keys())

    enriched_quotes = enrich_insurer_quotes(quotes, 20, 20)

    feature_keys.insert(0, "Total Premium")  # Add row at the top

    total_cols = 1 + len(insurer_list)
    column_width = Inches(9.0 / total_cols)

    for i, p in enumerate(doc.paragraphs):
        if placeholder in p.text:
            parent = p._element.getparent()
            idx = parent.index(p._element)
            parent.remove(p._element)
            tbl = doc.add_table(rows=1 + len(feature_keys), cols=total_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            tbl.autofit = True

            for col in range(total_cols):
                cell = tbl.cell(0, col)
                cell.width = column_width
                cell.text = "Common Policy Features" if col == 0 else insurer_list[col - 1]
                set_cell_background(cell, "FFFFFF")
                set_cell_bottom_border(cell)
                for p in cell.paragraphs:
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    for r in p.runs:
                        r.font.bold = True
                        r.font.size = Pt(11)
                        r.font.name = "Futura Bk BT"
                        r.font.color.rgb = RGBColor(0, 0, 0)

            for row_idx, key in enumerate(feature_keys, 1):
                color = "e9edf7" if row_idx % 2 == 1 else "FFFFFF"
                cell = tbl.cell(row_idx, 0)
                cell.width = column_width
                cell.text = key
                set_cell_background(cell, color)
                for p in cell.paragraphs:
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    for r in p.runs:
                        r.font.bold = True
                        r.font.size = Pt(9)
                        r.font.name = "Futura Bk BT"

                for col_idx, insurer in enumerate(insurer_list):
                    if key == "Total Premium":
                        val = enriched_quotes.get(insurer, {}).get("final_total", "-")
                        val = format_currency(val, 2) if is_number(val) else "-"
                    else:
                        val = quotes.get(insurer, {}).get("features", {}).get(key, "-")
                        if is_number(val):
                            val = format_currency(val, 0)
                    cell = tbl.cell(row_idx, col_idx + 1)
                    cell.width = column_width
                    cell.text = str(val) if val else "-"
                    set_cell_background(cell, color)
                    for p in cell.paragraphs:
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                        for r in p.runs:
                            r.font.size = Pt(9)
                            r.font.name = "Futura Bk BT"

            parent.insert(idx, tbl._element)
            break

# Modified insert_conditions_table with landscape page width support

def insert_conditions_table(doc, quotes):
    ensure_landscape_section(doc)
    placeholder = "{{conditions_table}}"
    total_cols = 2
    col_widths = [Inches(2.5), Inches(6.5)]

    for i, p in enumerate(doc.paragraphs):
        if placeholder in p.text:
            parent = p._element.getparent()
            idx = parent.index(p._element)
            parent.remove(p._element)
            tbl = doc.add_table(rows=1 + len(quotes), cols=total_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            tbl.autofit = True

            headers = ["Insurer", "Conditions / Endorsements"]
            for col in range(total_cols):
                cell = tbl.cell(0, col)
                cell.width = col_widths[col]
                cell.text = headers[col]
                set_cell_background(cell, "FFFFFF")
                set_cell_bottom_border(cell)
                for p in cell.paragraphs:
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    for r in p.runs:
                        r.font.bold = True
                        r.font.size = Pt(11)
                        r.font.name = "Futura Bk BT"
                        r.font.color.rgb = RGBColor(0, 0, 0)

            for row_idx, (insurer, quote) in enumerate(quotes.items(), 1):
                color = "e9edf7" if row_idx % 2 == 1 else "FFFFFF"
                values = [insurer, quote.get("conditions_or_endorsements", "-")]
                for col_idx, value in enumerate(values):
                    cell = tbl.cell(row_idx, col_idx)
                    cell.width = col_widths[col_idx]
                    cell.text = value if value else "-"
                    set_cell_background(cell, color)
                    for p in cell.paragraphs:
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                        for r in p.runs:
                            r.font.size = Pt(10)
                            r.font.name = "Futura Bk BT"

            parent.insert(idx, tbl._element)
            break

def generate_report(template_path, output_path, data, broker_fee_pct, commission_pct, associate_split,strata_manager):
    doc = Document(template_path)
    data["associate_split"] = associate_split
    data["strata_manager"] = strata_manager
    replace_dict = flatten_data_for_replace(data, broker_fee_pct, commission_pct,strata_manager)

    for p in doc.paragraphs:
        for key, value in replace_dict.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replace_dict.items():
                    if f"{{{{{key}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))

    enriched_quotes = enrich_insurer_quotes(data.get("Quotes", {}), broker_fee_pct, commission_pct, associate_split)
    recommended = find_recommended(enriched_quotes)

    insert_comparison_table(doc, data.get("Quotes", {}))
    insert_conditions_table(doc, data.get("Quotes", {}))
    insert_market_summary_table(doc, data.get("Quotes", {}), recommended.get("insurer"))

    doc.save(output_path)

if __name__ == "__main__":
    json_data = load_json("combined_quotes.json")
    broker_fee_pct = 20
    commission_pct = 20
    associate_split = 20  # Default value for testing
    strata_manager = "International Strata"
    template_path = "report_template.docx"
    output_path = "Clearlake Insurance Renewal Report 2025-2026.docx"
    generate_report(template_path, output_path, json_data, broker_fee_pct, commission_pct, associate_split,strata_manager)
    print("Report generated:", output_path)
