# extract.py

import pdfplumber
import json
from openai import OpenAI

openai = OpenAI()

with open("schema.json") as f:
    schema_json = json.load(f)

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

def extract_quote_data(text):
    messages = [
        {
            "role": "system",
            "content": f"""You are an assistant that extracts structured insurance quote data from unstructured PDF text.

Return the data in the following JSON structure:
{json.dumps(schema_json)}

Guidelines:
- If the quote doesn't mention a value, leave it blank or 0
- COPY THE SCHEMA's VALUE TYPES EXACTLY: if it is a blank string, your output for that entry must be a string. If it is a value PUT A VALUE.
- Extract excesses and limits if specified
- Only return the completed JSON object

Important to know when extracting:
- BSI is known as the building sum insured. It may also just be referred to as building. It is typically one of the first values. Should be in the hundreds of thousands / millions.
- For conditions, ignore anything related to time - such as how long the quotation is valid for. 
- IF GST is included in the underwriter / strata fee, remove gst (by dividing by 1.1) and just return the value without GST. Then put the GST in the underwriter gst section. Remember that computers struggle with floating point. I.e., 423.50 without GST is 385
- Any additional levies should be included in the ESL.
- ESL is commonly known as FSL
- Machinary breakdown is often called equipment breakdown
- Some insurers such as Insurance Investment Solutions will have many excesses. It is important to get every excess. They may have excesses on property, liability, voluntary workers, equipment, office bearers, and government audit and legal expenses
- Voluntary workers comp is also known as personal accident
- Additional benefits is the same as Extra benefits or additional limits
- FOR IIS in particular EXTRACT EVERY EXCESS (all the excesses you can expect are below)

Property Claims
Malicious Damage
Flood
Impact
New Construction
* All Standard Excess Claims (Discounted)
Burst Pipe &/or Resultant Water Damage
** Burst Flexi Pipe & Resultant Water Damage (Discounted)

Storm
Earthquake
Tropical Cyclone 

All Liability Claims
Claims involving Pool/spa
Claims involving Tennis Courts
Claims involving Playgrounds
Claims involving Gymnasium

All Voluntary Workers Claims

All Fidelity Excess Claims

All Water Chillers and Power Generators Claims
All Central AC Units Claims
All Small AC Units Claims
All Lift claims
All Other Equipment Breakdown Claims

Office Bearers Liability
Office Bearers Retroactive Date
Section 7 - Gov't Audit & Legal Expenses
Section 7a Taxation & Audit Excess
Section 7b Work Health Safety Excess
Section 7c Legal Expenses Excess
Section 7c Legal Expenses Contribution"""
        },
        {"role": "user", "content": text}
    ]

    response = openai.chat.completions.create(
        model="gpt-4o",  # or gpt-4-turbo if preferred
        messages=messages,
        temperature=0
    )
    return response.choices[0].message.content.strip()

def process_pdf(input_path, output_path):
    text = extract_text_from_pdf(input_path)
    result = extract_quote_data(text)
    with open(output_path, 'w') as f:
        f.write(result)
