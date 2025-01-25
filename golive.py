# Import necessary libraries
from flask import Flask, request, jsonify
import json
import requests
import base64
import extract_msg
import os
import pandas as pd
from docx import Document
import re
from datetime import datetime
from openai import AzureOpenAI
from io import BytesIO
from dotenv import load_dotenv
import asyncio
import aiohttp

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
async def extract_info_from_msg(file_path):
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

        async with aiohttp.ClientSession() as session:
            tasks = []
            for key, prompt in prompts.items():
                tasks.append(fetch_info(session, key, prompt, body))
            results = await asyncio.gather(*tasks)

        for key, result in results:
            info[key] = result

        if info["Completion Date"] == "Not Provided" or not re.search(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{4}\b', info["Completion Date"], re.IGNORECASE):
            if email_date:
                info["Completion Date"] = email_date.strftime("%B %Y")

        return info
    except Exception as e:
        raise

async def fetch_info(session, key, prompt, body):
    try:
        response = await client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"{prompt}\n\n{body}"}
            ]
        )
        return key, response.choices[0].message.content.strip()
    except Exception as api_error:
        return key, "Error extracting this field"

# Summarize extracted information
async def summarize_info(info):
    summary_prompts = {
        key: f"Summarize {key.lower()} briefly:" for key in [
            "Project Objectives", "Business Challenges", "Our Approach", "Value Created", "Measures of Success"
        ]
    }

    summarized_info = info.copy()
    async with aiohttp.ClientSession() as session:
        tasks = []
        for key, prompt in summary_prompts.items():
            tasks.append(fetch_summary(session, key, prompt, info[key]))
        results = await asyncio.gather(*tasks)

    for key, result in results:
        summarized_info[key] = result

    return summarized_info

async def fetch_summary(session, key, prompt, content):
    try:
        response = await client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"{prompt}\n\n{content}"}
            ]
        )
        return key, response.choices[0].message.content.strip()
    except Exception as e:
        return key, "Error summarizing this field"

# Flask routes
@app.route("/process-email", methods=["POST"])
async def process_email():
    try:
        data = request.json
        base64_content = data.get("$content")
        if not base64_content:
            return jsonify({"error": "No content provided"}), 400

        decoded_content = base64.b64decode(base64_content)
        email_stream = BytesIO(decoded_content)

        info = await extract_info_from_msg(email_stream)
        summarized_info = await summarize_info(info)
        response_data = {
            "extracted_info": info,
            "summarized_info": summarized_info
        }

        return jsonify(response_data)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/word", methods=["POST"])
async def worddocumentation():
    try:
        info = request.get_json()
        document_base64 = info.get("document")
        data = info.get("info")

        document_bytes = BytesIO(base64.b64decode(document_base64))
        modified_doc_bytes = create_summary_doc(document_bytes, [data])

        return jsonify({"doc_path": base64.b64encode(modified_doc_bytes.getvalue()).decode('utf-8')})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/excel", methods=["POST"])
async def exceldocumentation():
    try:
        data = request.get_json()
        excel_bytes = create_summary_excel([data])
        return jsonify({"excel_path": base64.b64encode(excel_bytes.getvalue()).decode('utf-8')})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
