from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict, List
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
from fastapi.responses import JSONResponse 
import json 
from docx.shared import Inches

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

class ExcelRequest(BaseModel):
    info: Dict[str, str] 
    
# Define request model
class WordRequest(BaseModel):
    data: str  # JSON string
    document: str  # Base64 encoded Word file

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
        "Project Title": "Provide the project title only (e.g., Real Estate Fund Portfolio Management):",         
        "Client Name": "Provide the client name (not Lionpoint) only (e.g., Bain Capital LP):",         
        "Use Case": "Provide the specific use cases only (e.g., Real Estate Fund Forecasting, Waterfall, Asset Management, Workforce Planning):",         
        "Completion Date": "Provide the completion date (Month and Year only, e.g., December 2021):",
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
                model = model_name,
                messages=payload["messages"]
            )
            return key, response.choices[0].message.content.strip()
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
            {"role": "system", "content": "Extract relevant information from the provided email and organize it into structured data to populate a word document template. The system will analyze the email content to identify key details, such as names, dates, or other specified information, and structure the output in a pre-determined format designed for seamless integration into a word document template."},
            {"role": "user", "content": f"{prompt}\n\n{content}"}
        ]
    }

    try:
        async with semaphore:
            response = await client.chat.completions.create(
                model = model_name,
                messages=payload["messages"]
            )
            return key, response.choices[0].message.content.strip()
    except Exception as e:
        return key, f"Error: {str(e)}"

# Generate Stream
async def stream_processor(response):
    async for chunk in response:
        if len(chunk.choices) > 0:
            delta = chunk.choices[0].delta
            if delta.content:
                yield delta.content 
                

async def add_heading_and_text(doc, heading, text, style=None):
    # Add the section heading
    doc.add_heading(heading, level=2)

    # Check if text is a non-empty string
    if isinstance(text, str) and text.strip():  # Check if text is non-empty
        paragraphs = text.split('\n')  # Split by new line

        for line in paragraphs:
            line = line.strip()  # Remove whitespace
            if line:  # Only proceed if the line is not empty
                # Check for bold text within ** **
                segments = line.split('**')
                paragraph = doc.add_paragraph()  # Create a new paragraph
                

                # Proper handling for bullet points or numbered items based on the first character
                if line.startswith('•') or line.startswith('-') :  # Check if it starts with a bullet
                    bullet_text = line[1:].strip()  # Remove the bullet character
                    bullet_paragraph = doc.add_paragraph(bullet_text)
                    bullet_paragraph.style = 'List Bullet'  # Set the style for bullet 
                    paragraph.paragraph_format.left_indent = Inches(0.50)  # Indent the numbered line
                elif line[0].isdigit():  # Check if it starts with a number
                    # For numbered lines, treat them as bold but add as separate paragraph 
                    for i, segment in enumerate(segments):
                        if i % 2 == 1:  # This means it is a bold text segment (between **)
                            paragraph.add_run(segment).bold = True  # Make it bold
                        else:  # Regular text
                            paragraph.add_run(segment)
                    paragraph.paragraph_format.left_indent = Inches(0.50)  # Indent the numbered line
                else:
                    # If it's plain text without bullets or numbers, just add it normally
                    paragraph.add_run(line)  # Add as normal text

    elif text:
        # If text is a single line without new lines
        if style:
            doc.add_paragraph(text, style=style)
        else:
            doc.add_paragraph(text)

async def create_summary_doc(existing_doc_bytes, all_data):
    # Decode the base64 document
    doc_content = base64.b64decode(existing_doc_bytes)

    # Create a BytesIO object to hold the document in memory
    doc_stream = BytesIO(doc_content)

    # Open the document
    doc = Document(doc_stream)

    # Process the provided project data
    for project in all_data:
        # Adding project title
        doc.add_heading(project['Project Title'], level=1)

        # Using the add_heading_and_text utility to add client information
        await add_heading_and_text(doc, 'Client Name:', project['Client Name'], 'Body Text')
        await add_heading_and_text(doc, 'Use Case:', project['Use Case'], 'Body Text')
        await add_heading_and_text(doc, 'Industry:', project['Industry'], 'Body Text')
        await add_heading_and_text(doc, 'Completion Date:', project['Completion Date'], 'Body Text')

        # Adding sections to the document
        await add_heading_and_text(doc, 'Project Objectives:', project['Project Objectives'], None)
        await add_heading_and_text(doc, 'Business Challenges:', project['Business Challenges'], None)
        await add_heading_and_text(doc, 'Our Approach:', project['Our Approach'], None)
        await add_heading_and_text(doc, 'Value Created:', project['Value Created'], None)
        await add_heading_and_text(doc, 'Measures of Success:', project['Measures of Success'], None)

    # Save the modified document to bytes
    updated_doc_bytes = BytesIO()
    doc.save(updated_doc_bytes)

    # Seek to the beginning to prepare for reading
    updated_doc_bytes.seek(0)

    # Convert the updated document to base64 for response
    encoded_doc = base64.b64encode(updated_doc_bytes.getvalue()).decode('utf-8')
    return encoded_doc

# API Endpoint for streaming
@app.post("/stream")
async def stream(prompt: Prompt):
    azure_open_ai_response = await client.chat.completions.create(
        model = model_name,
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
        # Extract document and project data
        document_base64 = request.document
        project_data_string = request.data
        
        # Parse the JSON string from 'data' into a list of dictionaries
        project_data = json.loads(project_data_string)
        
        if not document_base64 or not project_data:
            return JSONResponse(content={'error': 'Missing document or data'}, status_code=400)

        # Generate the updated document
        updated_doc_base64 = await create_summary_doc(document_base64, project_data)
        
        # Return the updated document as base64
        return JSONResponse(content={'modifiedDocument': updated_doc_base64}, status_code=200)
    except Exception as e:
        return JSONResponse(content={'error': str(e)}, status_code=500)
        
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
