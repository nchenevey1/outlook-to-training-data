"""
Author: Nicholas Chenevey
Date: 11/10/2024
This script processes email data and generates an email response template using a pre-trained model wwwith llama.cpp.
"""

from llama_cpp import Llama
import sys, os, re, io

# Function to suppress output
def suppress_output():
    sys.stdout = io.StringIO()  # Redirect stdout
    sys.stderr = io.StringIO()  # Redirect stderr

# Function to restore output
def restore_output():
    sys.stdout = sys.__stdout__  # Restore stdout
    sys.stderr = sys.__stderr__  # Restore stderr

# Function to ensure the byte string length is even
def ensure_even_length(hex_string):
    if len(hex_string) % 2 != 0:
        return hex_string[:-1]
    return hex_string

def extract_email_body(email_body):
    # Define patterns to match the start of quoted replies
    patterns = [
        r"On\s.+,\s.+\swrote:",              # Pattern for: On [Date], [Name] wrote:
        r"On\s.+,\s\d{4}.*<.*?>\s*wrote:",   # Handles variations with HTML-like tags in the email
        r"_{5,}\s*From:",                    # Pattern for lines that start with 5 or more underscores followed by "From:"
    ]
    
    # Combine the patterns into a single regex with OR condition
    combined_pattern = re.compile("|".join(patterns), re.MULTILINE | re.IGNORECASE)
    
    # Find the position of the first match, if any
    match = combined_pattern.search(email_body)
    
    if match:
        # Return everything before the match (i.e., the main email body)
        return email_body[:match.start()].strip()
    
    # If no match is found, return the full email body
    return email_body.strip()

def extract_text_after_subject(mainBody, extractAfter):
    # Escape any special characters in email_subject for regex
    escaped_subject = re.escape(extractAfter)
    
    # Pattern to find the subject in a case-insensitive manner and extract text that follows it
    pattern = re.compile(escaped_subject + r"(.+)", re.IGNORECASE | re.DOTALL)
    
    # Search for the pattern in main_body
    match = pattern.search(mainBody)
    
    # If a match is found, return the text after the subject; otherwise, return the full body
    return match.group(1).strip() if match else mainBody.strip()

# Suppress output
suppress_output()

# Load arguments from command line
# Ensure all byte strings have even length
emailBodyBytes = ensure_even_length(sys.argv[1][1:])
emailSubjectBytes = ensure_even_length(sys.argv[2][1:-1])
emailSenderBytes = ensure_even_length(sys.argv[3][1:-1])
emailTimeBytes = ensure_even_length(sys.argv[4][1:-1])

# Remove 'DA' from emailBodyBytes
emailBodyBytes = ''.join([emailBodyBytes[i:i+2] for i in range(0, len(emailBodyBytes), 2) if emailBodyBytes[i:i+2] != 'DA'])

# Decode byte strings to UTF-8
emailBody = bytes.fromhex(emailBodyBytes).decode('utf-8')
emailSubject = bytes.fromhex(emailSubjectBytes).decode('utf-8')
emailSender = bytes.fromhex(emailSenderBytes).decode('utf-8')
emailTime = bytes.fromhex(emailTimeBytes).decode('utf-8')

# Extract the main body of the email
mainBody = extract_email_body(emailBody)
mainBody = extract_text_after_subject(mainBody, emailSubject)

# Specify the path to the model file
model_path = "path/to/model/model.gguf"

## Instantiate model from downloaded file
llm = Llama(
    model_path=model_path,
    n_ctx=1000,  # Context length to use
    n_threads=8,  # Number of CPU threads to use
    n_gpu_layers=16  # Number of model layers to offload to GPU
)

token_count = 128

## Generation kwargs
generation_kwargs = {
    "max_tokens":token_count,
    "stop":["</s>"],
    "echo":False, # Echo the prompt in the output
    "top_k":1 # This is essentially greedy decoding, since the model will always return the highest-probability token. Set this value > 1 for sampling decoding
}

## Run inference
User_Name = "USER"
prompt = "Do not provide anything other than the reply body. Try to keep the output below " + str(token_count) + " total characters. Don't say anything after signing off as " + User_Name + ". You are " + User_Name + ", write a friendly email response template body to an email that was received and had the following parameters: From: '" + emailSender + "' , To: " + User_Name + ", Sent Date: '" + emailTime + "', Subject: '" + emailSubject + "', Body: '" + mainBody + "'"
res = llm(prompt, **generation_kwargs) # Res is a dictionary

## Unpack and the generated text from the LLM response dictionary and print it
email_body = res["choices"][0]["text"]

##### Optional Additional Processing #####

# Replace new lines with a placeholder string before printing
email_body = email_body.replace('\n', '!NEWLINE!')

# email_body = email_body.split(':', 1)[1]

# Remove leading newlines from email_body
email_body = email_body.lstrip('!NEWLINE!')

restore_output()
print(email_body)