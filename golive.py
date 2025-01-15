from flask import Flask, request, jsonify
import base64
import extract_msg
import os
import pandas as pd
from docx import Document
import re
from datetime import datetime
from openai import AzureOpenAI
from azure.core.credentials import AzureKeyCredential 
from io import BytesIO 
from dotenv import load_dotenv 
import logging 
# Set up logging
logging.basicConfig(level=logging.DEBUG)  # Logs to console for easier debugging
logger = logging.getLogger(__name__)


# Load .env only for local development (when not on Azure)
if os.getenv("FLASK_ENV") != "production":
    load_dotenv()  # This loads the .env file for local development

# Retrieve Azure environment variables from Azure App Service Application Settings
api_key = os.getenv("AZURE_OPENAI_API_key")
azure_endpoint = os.getenv("AZURE_OPENAI")

if not api_key or not azure_endpoint:
    raise EnvironmentError("AZURE_OPENAI_API_key or AZURE_OPENAI not set!")

# Set up environment variables  
client = AzureOpenAI(
    api_key= api_key,
    azure_endpoint= azure_endpoint,
    api_version="2024-08-01-preview",
)

# Set up Azure AI Project Client
# Set up OpenAI Client

app = Flask(__name__)

# Function to extract info from the .msg file
def extract_info_from_msg(file_path):
    msg = extract_msg.Message(file_path) 
     
    # Check if the email body is empty 
    email_date = msg.date
    email_subject = msg.subject 

    info = {
        "Project Title": "Not Provided",
        "Client Name": "Not Provided",
        "Use Case": "Not Provided",
        "Completion Date": "Not Provided",
        "Project Objectives": "Not Provided",
        "Business Challenges": "Not Provided",
        "Our Approach": "Not Provided",
        "Value Created": "Not Provided",
        "Measures of Success": "Not Provided",
        "Industry": "Not Provided"
    }

    body = re.sub(r'<[^>]+>', '', msg.body)       
    body = re.sub(r'\s+', ' ', body).strip()

    prompts = {
        "Project Title": "Extract the project title:",
        "Client Name": "Extract the client name (not Lionpoint):",
        "Use Case": "Extract the specific use case or objective of the project:",
        "Completion Date": "Extract the completion date (Month and Year):",
        "Project Objectives": "Extract the main objectives of the project:",
        "Business Challenges": "Extract the key business challenges faced by the client:",
        "Our Approach": "Extract the approach taken during the project:",
        "Value Created": "Extract the value created or outcomes achieved from the project:",
        "Measures of Success": "Extract the measures of success for the project:",
        "Industry": "Extract the industry related to the project:"
    }

    for key, prompt in prompts.items(): 
        try:  
            messages = [
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"{prompt}\n\n{body}"}
                ]
            response = client.chat.completions.create(
                model ="gpt-4o",
                messages=messages
            )
            
            info[key] = response.choices[0].message.content.strip()
            
            
        except Exception as api_error:
            print(f"OpenAI API call failed for '{key}': {api_error}")
            info[key] = "Error extracting this field"

    if info["Completion Date"] == "Not Provided" or not re.search(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{4}\b', info["Completion Date"], re.IGNORECASE):
        if email_date is not None:
            info["Completion Date"] = email_date.strftime("%B %Y")

    return info

# Function to summarize extracted information
def summarize_info(info):
    summary_prompts = {
        "Project Objectives": "Summarize the project objectives briefly:",
        "Business Challenges": "Summarize the business challenges faced briefly:",
        "Our Approach": "Summarize our approach briefly:",
        "Value Created": "Summarize the value created briefly:",
        "Measures of Success": "Summarize the measures of success briefly:"
    }

    summarized_info = info.copy()

    for key, prompt in summary_prompts.items():
        if key in info:
            messages = [
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"{prompt}\n\n{info[key]}"}
                ]
            response = client.chat.completions.create(
                model ="gpt-4o",
                messages=messages 
            )
            summarized_info[key] =  response.choices[0].message.content.strip()

    return summarized_info

# Function to create documentation
def create_summary_doc(existing_doc_bytes, all_data):
    # Load the existing document
    doc = Document(existing_doc_bytes)

    # Add new content
    for project in all_data:
        doc.add_heading(project['Project Title'], level=1)
        doc.add_paragraph(f"Client Name: {project['Client Name']}")
        doc.add_paragraph(f"Use Case: {project['Use Case']}")
        doc.add_paragraph(f"Industry: {project['Industry']}")
        doc.add_paragraph(f"Completion Date: {project['Completion Date']}")
        doc.add_paragraph("Project Objectives:", style='BodyText')
        doc.add_paragraph(project['Project Objectives'])
        doc.add_paragraph("Business Challenges:")
        doc.add_paragraph(project['Business Challenges'])
        doc.add_paragraph("Our Approach:")
        doc.add_paragraph(project['Our Approach'])
        doc.add_paragraph("Value Created:")
        doc.add_paragraph(project['Value Created'])
        doc.add_paragraph("Measures of Success:")
        doc.add_paragraph(project['Measures of Success'])
        doc.add_page_break()

    # Save the modified document to bytes
    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    return doc_bytes
    

def create_summary_excel(summarized_data):
    df = pd.DataFrame(summarized_data)
    excel_bytes = BytesIO()
    with pd.ExcelWriter(excel_bytes, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    excel_bytes.seek(0)
    

@app.route("/process-email", methods=["POST"])
def process_email(): 
    try:
        logger.debug("Processing /process-email route")
        data = request.json
        logger.debug(f"Received data: {data}")
        
        base64_content = data.get("$content", None)
        if not base64_content:
            logger.error("No content provided in request.")
            return jsonify({"error": "No content provided"}), 400

        decoded_content = base64.b64decode(base64_content)
        file_name = "temp_email.msg"
        with open(file_name, "wb") as file:
            file.write(decoded_content)

        info = extract_info_from_msg(file_name) 
        summarizedinfo = summarize_info(info) 
        response_data = {
                "extracted_info": info,
                "summarized_info": summarizedinfo
                }

        return jsonify(response_data)
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/word", methods=["POST"])
def worddocumentation():
    try:
        info = request.get_json() 
        document_base64 = info.get("document", None) 
        data = info.get("info", None) 
        
        all_data = pd.DataFrame([data])  
        
        # Decode the base64 document
        document_bytes = base64.b64decode(document_base64)
        existing_doc_bytes = BytesIO(document_bytes)

        modified_doc_bytes = create_summary_doc(existing_doc_bytes, all_data) 
        
        modified_document_base64 = base64.b64encode(modified_doc_bytes.getvalue()).decode('utf-8')

        return jsonify({
            "doc_path": modified_document_base64
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/excel", methods=["POST"])
def exceldocumentation():
    try:
        data = request.get_json()
        all_data = pd.DataFrame([data])

        if not all_data:
            return jsonify({"error": "No data provided"}), 400
        excel_path = create_summary_excel(all_data)

        return jsonify({
            "excel_path": excel_path
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
