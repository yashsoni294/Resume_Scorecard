# from fastapi import FastAPI, File, UploadFile
# import zipfile
# import os
# from io import BytesIO
# from PyPDF2 import PdfReader

# app = FastAPI()

# @app.post("/upload-zip/")
async def upload_zip(file: UploadFile = File(...)):
    # Ensure the uploaded file is a ZIP
    if not file.filename.endswith(".zip"):
        return {"error": "Uploaded file must be a ZIP archive."}

    # Read the uploaded ZIP file into memory
    zip_data = await file.read()
    zip_file = BytesIO(zip_data)
    
    # Extract the ZIP file
    extracted_files = []
    with zipfile.ZipFile(zip_file, 'r') as z:
        # Create a directory for extracted files (optional)
        extract_path = "extracted_files"
        os.makedirs(extract_path, exist_ok=True)
        
        # Extract files
        z.extractall(extract_path)
        extracted_files = z.namelist()

    # Read PDF files inside the ZIP
    pdf_contents = {}
    for file_name in extracted_files:
        if file_name.endswith(".pdf"):
            pdf_path = os.path.join(extract_path, file_name)
            with open(pdf_path, "rb") as pdf_file:
                reader = PdfReader(pdf_file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                pdf_contents[file_name] = text

    return {
        "message": "ZIP file processed successfully.",
        "pdf_contents": pdf_contents,  # Return PDF contents or process them further
    }


from fastapi import FastAPI, File, UploadFile
from io import BytesIO
import os
import time
from typing import List
from PyPDF2 import PdfReader
from docx import Document
import win32com.client as win32
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from PyPDF2 import PdfReader
from docx import Document
import win32com.client as win32
import os
import io
import zipfile
import rarfile  # Import rarfile library
import os
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import time
import tkinter as tk
from tkinter import filedialog
from dotenv import load_dotenv
import google.generativeai as genai
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import PromptTemplate
from datetime import datetime

# Load API Key
load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=API_KEY)


# Constants
TEMPLATES = {
    "job_description": """The text below is a job description:

            {job_description_text}

            Your task is to summarize the job description by focusing on the following areas:

            ### 1. **Role Requirements**
            - What are the primary responsibilities and duties expected in this role?
            - What specific experiences or industry backgrounds are preferred or required?

            ### 2. **Core Skills and Technical Expertise**
            - What technical and soft skills are required or preferred for this role?
            - Is there any emphasis on proficiency level or depth in these skills?

            ### 3. **Additional Requirements or Preferences**
            - Are there any other relevant requirements, such as language proficiency, travel, or work authorization?
            - Are there specific personality traits, work styles, or soft skills emphasized?

            Provide a concise summary based on these criteria, ensuring that no extraneous information is added. This summary will be used for evaluating candidate resumes. 
            """,
    "resume":""" 
            The text below is a resume:

            {resume_text}

            Your task is to summarize the resume by focusing on the following areas:

            ### 1. **Experience Relevance** 
            - Does the candidate have relevant role-specific and industry experience? 
            - How much experience does the candidate has and in which Domain ?

            ### 2. **Skills Alignment**
            - What core technical skills are demonstrated? 
            - How proficient is the candidate in these skills?

            ### 3. **Education & Certifications**
            - Does the candidate meet educational requirements? 
            - Are there any relevant certifications or additional learning?

            Provide a brief and focused summary based on these criteria. Also remember your given summary will be used for 
            evaluating the resume so do not add any extra information.
            """,
    "score": """
        Your task is to evaluate how well the provided resume aligns with the given job description by assessing three key factors: skills, experience, and education. Based on this evaluation, provide a final score between 0 and 100, representing the overall suitability of the candidate for the job.

        ### Inputs:
        - **Resume Text:**  
        {resume_text}

        - **Job Description Text:**  
        {job_description}

        ### Scoring Guidelines:
        Evaluate the resume against the job description across the following three dimensions. Provide a score for each dimension on a scale of 0 to 100:

        1. **Skills Alignment (0-100):**  
        Assess how well the skills mentioned in the resume match the skills required in the job description. Consider technical, soft, and specialized skills.

        2. **Experience Alignment (0-100):**  
        Compare the candidate work experience with the requirements outlined in the job description. Focus on factors such as relevance, industry alignment, and duration of experience.

        3. **Education Alignment (0-100):**  
        Evaluate the candidate's educational background against the academic qualifications specified in the job description. Consider the field of study, degree level, and institution if applicable.

        ### Final Score Calculation:
        Calculate the weighted average of the three scores based on the following weights:
            - Experience: 40%
            - Skills: 40%
            - Education: 20%

        Formula:  
        Final Score = (Experience Score * 0.4) + (Skills Score * 0.4)  + (Education Score * 0.2)

        Round the result to the nearest whole number.

        ### Output:
        Provide only the final average score as a single number (0-100) with no additional text or explanation.
    """
    ,
}


app = FastAPI()

def get_conversation(template: str):
    """
    Initializes a conversation using a template and LangChain's LLM integration.
    """
    llm = ChatGoogleGenerativeAI(
        model="gemini-1.5-flash",
        temperature=0.1,
        max_tokens=None,
        timeout=None,
        max_retries=3
    )
    prompt = PromptTemplate.from_template(template)
    return prompt | llm

def read_pdf(file: BytesIO) -> str:
    """Extract text from a PDF file."""
    try:
        pdf_reader = PdfReader(file)
        extracted_text = "".join([page.extract_text() for page in pdf_reader.pages])
        return extracted_text.strip()
    except Exception as e:
        return f"Error processing PDF: {str(e)}"

def read_docx(file: BytesIO) -> str:
    """Extract text from a DOCX file."""
    try:
        document = Document(file)
        extracted_text = "\n".join([paragraph.text for paragraph in document.paragraphs])
        return extracted_text.strip()
    except Exception as e:
        return f"Error processing DOCX: {str(e)}"

def read_doc(file_path: str) -> str:
    """Extract text from a DOC file using COM automation (Windows only)."""
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(file_path))
        extracted_text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        return extracted_text.strip()
    except Exception as e:
        return f"Error processing DOC file: {str(e)}"

def read_txt(file: BytesIO) -> str:
    """Extract text from a plain text file."""
    try:
        return file.read().decode("utf-8").strip()
    except Exception as e:
        return f"Error processing TXT file: {str(e)}"

def process_file(file_name: str, file_content: bytes) -> str:
    """
    Determine the file type and extract text based on its format.
    """
    file_extension = file_name.split(".")[-1].lower()
    file_stream = BytesIO(file_content)

    if file_extension == "pdf":
        return read_pdf(file_stream)
    elif file_extension == "docx":
        return read_docx(file_stream)
    elif file_extension == "doc":
        temp_file_path = f"temp_{file_name}"
        with open(temp_file_path, "wb") as temp_file:
            temp_file.write(file_content)
        extracted_text = read_doc(temp_file_path)
        os.remove(temp_file_path)
        return extracted_text
    elif file_extension == "txt":
        return read_txt(file_stream)
    elif file_extension == "zip":
        pass
    else:
        return f"Unsupported file type: {file_name}"

async def process_files(files: List[UploadFile], job_description: str) -> dict:
    """
    Process uploaded files and extract text.
    """
    response_data = {}

    conversation_jd = get_conversation(TEMPLATES["job_description"])
    jd_response = conversation_jd.invoke({"job_description_text": job_description})
    processed_jd = jd_response.content
    print(processed_jd)

    for file in files:
        try:
            file_content = await file.read()
            extracted_text = process_file(file.filename, file_content)
            response_data[file.filename] = {"content": extracted_text}
        except Exception as e:
            response_data[file.filename] = {"error": str(e)}

    return response_data

async def score_resumes(response_data: dict, job_description: str) -> dict:
    """
    Score resumes based on job description.
    """
    conversation_resume = get_conversation(TEMPLATES["resume"])
    conversation_score = get_conversation(TEMPLATES["score"])

    for filename, file_data in response_data.items():
        resume_response = conversation_resume.invoke({"resume_text": file_data["content"]})
        time.sleep(3)  # Introducing delay to avoid rate limiting

        file_data["key_feature"] = resume_response.content

        score_response = conversation_score.invoke({
            "resume_text": resume_response.content,
            "job_description": job_description
        })
        time.sleep(3)  # Introducing delay to avoid rate limiting

        file_data["score"] = score_response.content
        print(f"{filename}: {score_response.content}")

    return response_data

@app.post("/upload-files/")
async def upload_files(job_description: str, files: List[UploadFile] = File(...)):
    """
    Handle file uploads and process resumes based on the job description.
    """
    try:
        # Step 1: Process the uploaded files
        response_data = await process_files(files, job_description)

        # Step 2: Score the resumes
        scored_data = await score_resumes(response_data, job_description)

        return scored_data

    except Exception as e:
        return {"error": f"Error processing files: {str(e)}"}

