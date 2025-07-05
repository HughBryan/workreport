import pdfplumber
import json
import re
import os
from openai import OpenAI
from dotenv import load_dotenv
import sys
from datetime import datetime
from openai import AzureOpenAI  

load_dotenv()

# --- Azure OpenAI Credentials ---
azure_api_key = os.getenv("AZURE_OPENAI_API_KEY")
azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")
azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION")

# Instantiate Azure OpenAI client
openai = AzureOpenAI(
    api_key=azure_api_key,
    azure_endpoint=azure_endpoint,
    api_version=azure_api_version,
)


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

def extract_quote_data(text,longitude_option):
    messages = [
        {
            "role": "system",
            "content": f"""You are an assistant that extracts structured insurance quote data from unstructured PDF text.

For each quote I give you, I want you to follow the below quote-schema JSON configuration. We are going to fill out all the information in the below JSON with the information from the quote pdf I provide. Most values of the 'features' section are either a dollar amount, or "included" or "Not Included". Try to stick to that. 


{json.dumps(quote_schema)}

Here is my current JSON file, it contains all the current information from previous quotes. I want you to update any of the general information, but once we have extracted all the data about our current quote, return the current JSON file, with the new quote information we have extracted. Return ONLY the updated JSON file.

{json.dumps(main_schema)}


 
You will be given insruance quotations from multiple insurers. This may be included but not limited to the following: 
CHU (also known as QBE), 
Flex (also known as: CHU iSaver / CHUISAVER), 
SUU, 
Hutch, 
Axis, 
Rubix, 
BARN, 
Longitude, 
QUS, 
SCI,  
IIS (also known as: Insurance Investment Solutions).

When there is another alias for the insurer, use the abbreviated main name I have provided above.


Guidelines:
- If the quote doesn't mention a value, leave it blank or 0
- COPY THE SCHEMA's VALUE TYPES EXACTLY: if it is a blank string, your output for that entry must be a string. If it is a value PUT A VALUE. For example, Voltunary worker is a 0 in the schema - indicating that it needs the exact value the insurer sets.
- Extract excesses and limits if specified
- Only return the completed JSON object
- IF there is a terrorism levy, add it to base premium
- Longitude offers a 'current option' and a 'suggested option', use the {longitude_option}



Rules for Additional and extra benefits:
- Additional benefits is the same as Extra benefits or additional limits. Additional benefits should NOT include any other features already stated! For example, do not include catastrophe cover, lot owners improvements, loss of rent etc. in this category.
- "Additional and extra benefits commonly includes: taxation and audit costs, workpalce health and safety breaches, and legal defence expenses. 
- LIST the dollar value associated with each additional / extract benefit.

Rules for Building Sum Insured:
- BSI is known as the building sum insured. 
- It may also just be referred to as building. 
- If BSI / Building Sum Insured is not in the document, leave it as 0.
-  Do not get this confused with the liability to others 'sum insured'.
- If you are reading a 'quote response form', there is a likely chance that you will not have the BSI/building sum insured and should leave BSI 0.


Important to know when extracting:
- For conditions, ignore anything related to time - such as how long the quotation is valid for. 
- IF GST is included in the underwriter / strata fee, remove gst (by dividing by 1.1) and just return the value without GST. Then put the GST in the underwriter gst section. Remember that computers struggle with floating point. I.e., 423.50 without GST is 385
- Any additional levies should be included in the ESL.
- ESL is commonly known as FSL
- Machinary breakdown is often called equipment breakdown
- Some insurers such as Insurance Investment Solutions will have many excesses. It is important to get every excess. They may have excesses on property, liability, voluntary workers, equipment, office bearers, and government audit and legal expenses
- Voluntary workers comp is also known as personal accident and is typically $200,000
- For common contents if its not specified, just say 'Included in BSI'.  
- Fidelity is also known as: loss of funds, theft of funds.
- Always use the 'Insurer Alternative value' for all features such as common contents.
- Public Liability (also known as liability to others) is always 10,20,30,40 or 50 million dollars.


- YOU MUST list the COST of EVERY additional excess(es)
- Extract the value listed for “Base Premium” in the main premium breakdown, not any subtotals or split lines.
- When extracting the strata plan, just remove SP from the start. I.e., SP1234 extracted is 1234


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


def process_folder(folder_path, output_path, longitude_option = "Current Option", log_callback=None):
    import copy
    master = copy.deepcopy(main_schema)
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            text = extract_text_from_pdf(os.path.join(folder_path, filename))
            quote_json = extract_quote_data(text,longitude_option)
            update_master_json(master, quote_json)
            if log_callback:
                log_callback(f"{filename} has been completed")  # <-- This will print to GUI log
            else:
                print(f"{filename} has been completed")
    master["general_info"]["current_date"] = (datetime.now().strftime("%d/%m/%Y"))
    with open(output_path, "w") as f:
        json.dump(master, f, indent=2)


# Optional: keep your single-file processor
def process_pdf(input_path, output_path):
    text = extract_text_from_pdf(input_path)
    result = extract_quote_data(text,"current insurer")
    with open(output_path, 'w') as f:
        json.dump(result, f, indent=2)

if __name__ == "__main__":
    print(process_pdf("Quote.pdf","output.json"))
