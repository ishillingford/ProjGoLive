from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict
import base64
import re
import asyncio
import aiohttp
from io import BytesIO
from datetime import datetime
import extract_msg
import os
from dotenv import load_dotenv
import concurrent.futures

# Load environment variables
if os.getenv("FLASK_ENV") != "production":
    load_dotenv()

api_key = os.getenv("AZURE_OPENAI_API_key")
azure_endpoint = os.getenv("AZURE_OPENAI")
model_name = "gpt-4o-mini"
if not api_key or not azure_endpoint:
    raise EnvironmentError("AZURE_OPENAI_API_key or AZURE_OPENAI not set!")

# FastAPI app initialization
app = FastAPI()

# Semaphore to limit concurrent API requests
semaphore = asyncio.Semaphore(5)

# Request model
class EmailRequest(BaseModel):
    content: str  # Base64-encoded .msg file

# Extract information from .msg file
def sync_extract_msg(file_path):
    """Extracts metadata and body from a .msg file."""
    msg = extract_msg.Message(file_path)
    email_date = msg.date
    body = re.sub(r'<[^>]+>', '', msg.body)  # Remove HTML tags
    body = re.sub(r'\s+', ' ', body).strip()  # Normalize whitespace
    return email_date, body

async def extract_info_from_msg(file_path):
    """Extracts structured information from the .msg file."""
    # Offload blocking task to thread pool
    loop = asyncio.get_event_loop()
    with concurrent.futures.ThreadPoolExecutor() as pool:
        email_date, body = await loop.run_in_executor(pool, sync_extract_msg, file_path)

    # Prompts for information extraction
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

    # Collect results using async tasks
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_info(session, key, prompt, body) for key, prompt in prompts.items()]
        results = await asyncio.gather(*tasks)

    # Build the result dictionary
    info = {key: result for key, result in results}

    # Handle completion date fallback
    if info["Completion Date"] == "Not Provided" or not re.search(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{4}\b',
        info["Completion Date"],
        re.IGNORECASE
    ):
        if email_date:
            info["Completion Date"] = email_date.strftime("%B %Y")

    return info

async def fetch_info(session, key, prompt, body):
    """Fetches specific information using Azure OpenAI."""
    url = f"{azure_endpoint}/openai/deployments/{model_name}/chat/completions?api-version=2024-08-01-preview"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"{prompt}\n\n{body}"}
        ]
    }

    try:
        async with semaphore, session.post(url, json=payload, headers=headers) as response:
            if response.status == 200:
                data = await response.json()
                return key, data['choices'][0]['message']['content'].strip()
            else:
                return key, f"Error: {response.status}"
    except Exception as e:
        return key, f"Error: {str(e)}"

# Summarize extracted information
async def summarize_info(info):
    """Summarizes key information fields."""
    summary_prompts = {
        key: f"Summarize {key.lower()} briefly:" for key in [
            "Project Objectives", "Business Challenges", "Our Approach", "Value Created", "Measures of Success"
        ]
    }

    async with aiohttp.ClientSession() as session:
        tasks = [fetch_summary(session, key, prompt, info[key]) for key, prompt in summary_prompts.items()]
        results = await asyncio.gather(*tasks)

    summarized_info = info.copy()
    for key, result in results:
        summarized_info[key] = result

    return summarized_info

async def fetch_summary(session, key, prompt, content):
    """Fetches summarized content using Azure OpenAI."""
    url = f"{azure_endpoint}/openai/deployments/{model_name}/chat/completions?api-version=2024-08-01-preview"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"{prompt}\n\n{content}"}
        ]
    }

    try:
        async with semaphore, session.post(url, json=payload, headers=headers) as response:
            if response.status == 200:
                data = await response.json()
                return key, data['choices'][0]['message']['content'].strip()
            else:
                return key, f"Error: {response.status}"
    except Exception as e:
        return key, f"Error: {str(e)}"

# API Endpoint
@app.post("/process-email")
async def process_email(request: EmailRequest):
    try:
        base64_content = request.content
        if not base64_content:
            raise HTTPException(status_code=400, detail="No content provided.")

        # Decode the .msg file
        decoded_content = base64.b64decode(base64_content)
        email_stream = BytesIO(decoded_content)

        # Extract and summarize information
        extracted_info = await extract_info_from_msg(email_stream)
        summarized_info = await summarize_info(extracted_info)

        return {
            "extracted_info": extracted_info,
            "summarized_info": summarized_info
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

