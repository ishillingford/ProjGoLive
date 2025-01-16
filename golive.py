# Import necessary libraries
import os
import re
import logging
import base64
from io import BytesIO
import pandas as pd
from flask import Flask, request, jsonify
from flask.logging import default_handler
from docx import Document
from azure.ai.openai import AzureOpenAI
from dotenv import load_dotenv
import extract_msg

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
logger.addHandler(default_handler)

# Load environment variables
if os.getenv("FLASK_ENV") != "production":
    load_dotenv()

api_key = os.getenv("AZURE_OPENAI_API_key")
azure_endpoint = os.getenv("AZURE_OPENAI")
if not api_key or not azure_endpoint:
    raise EnvironmentError("AZURE_OPENAI_API_key or AZURE_OPENAI not set!")

# Set up Azure OpenAI client
client = AzureOpenAI(
    api_key=api_key,
    azure_endpoint=azure_endpoint,
    api_version="2024-08-01-preview",
)

app = Flask(__name__)

# Extract information from .msg file
def extract_info_from_msg(file_path):
    try:
        msg = extract_msg.Message(file_path)
        email_date = msg.date
        email_subject = msg.subject

        info = {key: "Not Provided" for key in [
            "Project Title", "Client Name", "Use Case", "Completion Date",
            "Project Objectives", "Business Challenges", "Our Approach",
            "Value Created", "Measures of Success", "Industry"
        ]}

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
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant."},
                        {"role": "user", "content": f"{prompt}\n\n{body}"}
                    ]
                )
                info[key] = response.choices[0].message.content.strip()
            except Exception as api_error:
                logger.error(f"OpenAI API call failed for '{key}': {api_error}")
                info[key] = "Error extracting this field"

        if info["Completion Date"] == "Not Provided" or not re.search(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{4}\b', info["Completion Date"], re.IGNORECASE):
            if email_date:
                info["Completion Date"] = email_date.strftime("%B %Y")

        return info
    except Exception as e:
        logger.error(f"Error extracting info from .msg: {e}")
        raise

# Summarize extracted information
def summarize_info(info):
    summary_prompts = {
        key: f"Summarize {key.lower()} briefly:" for key in [
            "Project Objectives", "Business Challenges", "Our Approach", "Value Created", "Measures of Success"
        ]
    }

    summarized_info = info.copy()
    for key, prompt in summary_prompts.items():
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": f"{prompt}\n\n{info[key]}"}
                ]
            )
            summarized_info[key] = response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"Error summarizing '{key}': {e}")
            summarized_info[key] = "Error summarizing this field"
    return summarized_info

# Flask routes
@app.route("/process-email", methods=["POST"])
def process_email():
    try:
        data = request.json
        base64_content = data.get("$content")
        if not base64_content:
            return jsonify({"error": "No content provided"}), 400

        decoded_content = base64.b64decode(base64_content)
        file_name = "temp_email.msg"
        with open(file_name, "wb") as file:
            file.write(decoded_content)

        info = extract_info_from_msg(file_name)
        summarized_info = summarize_info(info)
        return jsonify({"extracted_info": info, "summarized_info": summarized_info})
    except Exception as e:
        logger.error(f"Error processing /process-email: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/word", methods=["POST"])
def worddocumentation():
    try:
        info = request.get_json()
        document_base64 = info.get("document")
        data = info.get("info")

        document_bytes = BytesIO(base64.b64decode(document_base64))
        modified_doc_bytes = create_summary_doc(document_bytes, [data])

        return jsonify({"doc_path": base64.b64encode(modified_doc_bytes.getvalue()).decode('utf-8')})
    except Exception as e:
        logger.error(f"Error creating Word documentation: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/excel", methods=["POST"])
def exceldocumentation():
    try:
        data = request.get_json()
        excel_bytes = create_summary_excel([data])
        return jsonify({"excel_path": base64.b64encode(excel_bytes.getvalue()).decode('utf-8')})
    except Exception as e:
        logger.error(f"Error creating Excel documentation: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

