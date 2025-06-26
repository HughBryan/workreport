import pdfplumber
import json
import re
import os
from openai import OpenAI
from dotenv import load_dotenv
import sys
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")
openai = OpenAI()



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


with open(resource_path("main_schema.json")) as f:
    main_schema = json.load(f)

with open(resource_path("quote_schema.json")) as f:
    quote_schema = json.load(f)


def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

def extract_quote_data(text):
    messages = [
        {
            "role": "system",
            "content": f"""You are an assistant that extracts structured insurance quote data from unstructured PDF text.

For each quote I give you, I want you to follow the below quote-schema JSON configuration. We are going to fill out all the information in the below JSON with the information from the quote pdf I provide. Most values of the 'features' section are either a dollar amount, or "included" or "Not Included". Try to stick to that. 

For common contents if its not specified, just say 'Included in BSI'.  

Always use the 'Insurer Alternative value' for all features such as common contents.

Paint & wallpaper is always "Included".
{json.dumps(quote_schema)}

Here is my current JSON file, it contains all the current information from previous quotes. I want you to update any of the general information, but once we have extracted all the data about our current quote, return the current JSON file, with the new quote information we have extracted. Return ONLY the updated JSON file.

{json.dumps(main_schema)}


 
You will be given insruance quotations from multiple insurers, including but not limited to: CHU (also known as QBE), Flex (also known as CHUISAVER), SUU, Hutch, Axis, Rubix, BARN, Longitude, QUS, SCI,  IIS (insurance investment solutions) etc. Try use these names as the insurers name if they match.

Guidelines:
- If the quote doesn't mention a value, leave it blank or 0
- COPY THE SCHEMA's VALUE TYPES EXACTLY: if it is a blank string, your output for that entry must be a string. If it is a value PUT A VALUE. For example, Voltunary worker is a 0 in the schema - indicating that it needs the exact value the insurer sets.
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
    content = response.choices[0].message.content.strip()

    # Remove ```json ... ``` or ``` ... ``` wrappers if present
    match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", content, re.DOTALL)
    if match:
        content = match.group(1).strip()
    return json.loads(content)

def update_master_json(master, new):
    # Update general_info
    if "general_info" in new:
        for k, v in new["general_info"].items():
            if v not in ("", 0, None):
                master["general_info"][k] = v
    # Update the correct insurer in Quotes
    if "Quotes" in new:
        for insurer, info in new["Quotes"].items():
            # If this insurer is not in master, add it
            if insurer not in master["Quotes"]:
                master["Quotes"][insurer] = info
                continue  # Already added all info, skip to next insurer

            # Otherwise update fields as before
            for field, value in info.items():
                if isinstance(value, dict):
                    if field not in master["Quotes"][insurer]:
                        master["Quotes"][insurer][field] = value
                        continue
                    for subfield, subvalue in value.items():
                        if subvalue not in ("", 0, None):
                            master["Quotes"][insurer][field][subfield] = subvalue
                else:
                    if value not in ("", 0, None):
                        master["Quotes"][insurer][field] = value


def process_folder(folder_path, output_path):
    import copy
    master = copy.deepcopy(main_schema)
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            text = extract_text_from_pdf(os.path.join(folder_path, filename))
            quote_json = extract_quote_data(text)
            update_master_json(master, quote_json)
    with open(output_path, "w") as f:
        json.dump(master, f, indent=2)

# Optional: keep your single-file processor
def process_pdf(input_path, output_path):
    text = extract_text_from_pdf(input_path)
    result = extract_quote_data(text)
    with open(output_path, 'w') as f:
        json.dump(result, f, indent=2)
