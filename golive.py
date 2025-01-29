from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict
import base64
import re
import asyncio
from io import BytesIO
from datetime import datetime
import extract_msg
import os
from dotenv import load_dotenv
import concurrent.futures
import pandas as pd
from docx import Document
from openai import AsyncAzureOpenAI
from starlette.responses import StreamingResponse

# Load environment variables
if os.getenv("FLASK_ENV") != "production":
    load_dotenv()

api_key = os.getenv("AZURE_OPENAI_API_key")
azure_endpoint = os.getenv("AZURE_OPENAI")
model_name = "gpt-4o-mini"
if not api_key or not azure_endpoint:
    raise EnvironmentError("AZURE_OPENAI_API_key or AZURE_OPENAI not set!")

# Initialize AsyncAzureOpenAI client
client = AsyncAzureOpenAI(
    api_key=api_key,
    azure_endpoint=azure_endpoint,
    api_version="2024-08-01-preview",
)

# FastAPI app initialization
app = FastAPI()

# Semaphore to limit concurrent API requests
semaphore = asyncio.Semaphore(5)

# Request models
class EmailRequest(BaseModel):
    content: str  # Base64-encoded .msg file

class WordRequest(BaseModel):
    document: str  # Base64-encoded Word document
    info: Dict[str, str]

class ExcelRequest(BaseModel):
    info: Dict[str, str]

class Prompt(BaseModel):
    input: str

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
    tasks = [fetch_info(key, prompt, body) for key, prompt in prompts.items()]
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

async def fetch_info(key, prompt, body):
    """Fetches specific information using Azure OpenAI."""
    payload = {
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"{prompt}\n\n{body}"}
        ]
    }

    try:
        async with semaphore:
            response = await client.chat.completions.create(
                deployment_id=model_name,
                messages=payload["messages"]
            )
            return key, response['choices'][0]['message']['content'].strip()
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

    tasks = [fetch_summary(key, prompt, info[key]) for key, prompt in summary_prompts.items()]
    results = await asyncio.gather(*tasks)

    summarized_info = info.copy()
    for key, result in results:
        summarized_info[key] = result

    return summarized_info

async def fetch_summary(key, prompt, content):
    """Fetches summarized content using Azure OpenAI."""
    payload = {
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"{prompt}\n\n{content}"}
        ]
    }

    try:
        async with semaphore:
            response = await client.chat.completions.create(
                deployment_id=model_name,
                messages=payload["messages"]
            )
            return key, response['choices'][0]['message']['content'].strip()
    except Exception as e:
        return key, f"Error: {str(e)}"

# Generate Stream
async def stream_processor(response):
    async for chunk in response:
        if len(chunk.choices) > 0:
            delta = chunk.choices[0].delta
            if delta.content:
                yield delta.content

# API Endpoint for streaming
@app.post("/stream")
async def stream(prompt: Prompt):
    azure_open_ai_response = await client.chat_completions.create(
        deployment_id=model_name,
        messages=[{"role": "user", "content": prompt.input}],
        stream=True
    )

    return StreamingResponse(stream_processor(azure_open_ai_response), media_type="text/event-stream")

# API Endpoint for processing email
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

# API Endpoint for Word documentation
@app.post("/word")
async def word_documentation(request: WordRequest):
    try:
        document_base64 = request.document
        data = request.info

        document_bytes = BytesIO(base64.b64decode(document_base64))
        doc = Document(document_bytes)

        # Example: Add summary content to the Word document
        doc.add_heading("Summary", level=1)
        for key, value in data.items():
            doc.add_heading(key, level=2)
            doc.add_paragraph(value)

        modified_doc_bytes = BytesIO()
        doc.save(modified_doc_bytes)

        return {"document": base64.b64encode(modified_doc_bytes.getvalue()).decode('utf-8')}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# API Endpoint for Excel documentation
@app.post("/excel")
async def excel_documentation(request: ExcelRequest):
    try:
        data = request.info
        df = pd.DataFrame([data])

        modified_excel_bytes = BytesIO()
        df.to_excel(modified_excel_bytes, index=False)

        return {"excel": base64.b64encode(modified_excel_bytes.getvalue()).decode('utf-8')}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
