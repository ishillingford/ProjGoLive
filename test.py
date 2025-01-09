import json
import requests
import base64

# Flask app endpoints
url = "http://127.0.0.1:5000/process-email"  # Update to your local or Azure endpoint 
url1 = "http://127.0.0.1:5000/word"  # Update to your local or Azure endpoint

# Load input data from the file
with open('input_data.json', 'r') as file:
    data = json.load(file)

# Extract only the $content field
content = data.get("$content")

# Create the payload to send in the POST request
payload = {
    "document": content
}

# Send POST request to process-email endpoint
headers = {
    'Content-Type': 'application/json'
}
response = requests.post(url, json=payload, headers=headers)

# Print the response from process-email endpoint
print("Status Code (process-email):", response.status_code)
print("Response Body (process-email):", response.text)

# Assuming the response from process-email contains the array data needed for the word endpoint
array_data = response.json().get('extracted_info')  # Adjust based on actual response structure

# Create the payload for the word endpoint
word_payload = {
    "document": content,  # Assuming the document content is the same
    "info": array_data
}

# Send POST request to word endpoint
word_response = requests.post(url1, json=word_payload, headers=headers)

# Print the response from word endpoint
print("Status Code (word):", word_response.status_code)
print("Response Body (word):", word_response.text)

# Save the modified document returned from the word endpoint
if word_response.status_code == 200:
    modified_document_base64 = word_response.json().get('document')
    with open('modified_document.docx', 'wb') as file:
        file.write(base64.b64decode(modified_document_base64))
    print("Modified document saved as 'modified_document.docx'")
else:
    print("Failed to process the document")