from docx import Document
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

def format_currency(value):
    try:
        value = round(float(value), 2)
        return "${:,.2f}".format(value)
    except Exception:
        return str(value)

def calculate_broker_fee(base, broker_fee_pct, commission_pct, commission_without_gst):
    # Check the commission WITHOUT gst field for 0/"0"/""/None
    if commission_without_gst in ("", 0, None, "0"):
        percent = broker_fee_pct + commission_pct
    else:
        percent = broker_fee_pct
    return round(base * (percent/100.0), 2)

def enrich_insurer_quotes(quotes_dict, broker_fee_pct, commission_pct):
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
        final_total = round(total + broker_fee + broker_gst, 2)

        enriched_quote = dict(quote)
        enriched_quote["broker_fee"] = broker_fee
        enriched_quote["broker_gst"] = broker_gst
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

def flatten_data_for_replace(data, broker_fee_pct, commission_pct):
    flat = {}
    for k, v in data.get("general_info", {}).items():
        flat[k] = v
    enriched_quotes = enrich_insurer_quotes(data.get("Quotes", {}), broker_fee_pct, commission_pct)
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
    return flat

def generate_report(template_path, output_path, data, broker_fee_pct, commission_pct):
    doc = Document(template_path)
    replace_dict = flatten_data_for_replace(data, broker_fee_pct, commission_pct)

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
    doc.save(output_path)

# Example usage (remove in production):
if __name__ == "__main__":
    json_data = load_json("combined_quotes.json")
    broker_fee_pct = 20
    commission_pct = 20
    template_path = "report_template.docx"
    output_path =  "Clearlake Insurance Renewal Report 2025-2026.docx"
    generate_report(template_path, output_path, json_data, broker_fee_pct, commission_pct)
    print("Report generated:", output_path)
