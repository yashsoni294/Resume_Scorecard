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

def get_conversation(template):
    """
    Initializes a conversation using a template and LangChain's LLM integration.
    Args:
        template (str): Template content for the conversation.
    Returns:
        Callable: Configured conversation object.
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

def read_pdf(file: io.BytesIO):
    """Extract text from a PDF file."""
    pdf_reader = PdfReader(file)
    extracted_text = ""
    for page in pdf_reader.pages:
        extracted_text += page.extract_text()
    return extracted_text.strip()

def read_docx(file: io.BytesIO):
    """Extract text from a DOCX file."""
    document = Document(file)
    extracted_text = ""
    for paragraph in document.paragraphs:
        extracted_text += paragraph.text + "\n"
    return extracted_text.strip()

def read_doc(file_path: str):
    """
    Extract text from a DOC file using COM automation (Windows only).
    """
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

def read_txt(file: io.BytesIO):
    """Extract text from a plain text file."""
    try:
        contents = file.read()
        return contents.decode("utf-8").strip()
    except Exception as e:
        return f"Error processing TXT file: {str(e)}"

def process_file(file_name: str, file_content: bytes):
    """
    Determine the file type and extract text based on its format.
    """
    try:
        file_extension = file_name.split(".")[-1].lower()
        file_stream = io.BytesIO(file_content)

        if file_extension == "pdf":
            return read_pdf(file_stream)
        elif file_extension == "docx":
            return read_docx(file_stream)
        elif file_extension == "doc":
            temp_file_path = f"temp_{file_name}"
            with open(temp_file_path, "wb") as temp_file:
                temp_file.write(file_content)
            text = read_doc(temp_file_path)
            os.remove(temp_file_path)
            return text
        elif file_extension == "txt":
            return read_txt(file_stream)
        else:
            return f"Unsupported file type: {file_name}"
    except Exception as e:
        return f"Error processing file {file_name}: {str(e)}"


@app.post("/upload-files/")
async def upload_files(job_description: str, files: list[UploadFile] = File(...)):
    response_data = {}

    conversation_jd = get_conversation(TEMPLATES["job_description"])
    jd_response = conversation_jd.invoke({"job_description_text": job_description})
    processed_jd = jd_response.content
    print(processed_jd)

    for file in files:
        try:
            # Fallback to file extension if MIME type is not reliable
            file_extension = file.filename.split(".")[-1].lower()

            if file.content_type == "application/zip" or file_extension == "zip":
                pass
                # response_data[file.filename] = process_zip(file)
            elif file.content_type in ["application/x-rar-compressed", "application/vnd.rar"] or file_extension == "rar":
                # response_data[file.filename] = process_rar(file)
                pass
            elif file.content_type == "application/pdf" or file_extension == "pdf":
                extracted_text = read_pdf(io.BytesIO(file.file.read()))
                response_data[file.filename] = {"content": extracted_text}
            elif file.content_type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"] or file_extension in ["docx", "doc"]:
                if file_extension == "docx":
                    extracted_text = read_docx(io.BytesIO(file.file.read()))
                elif file_extension == "doc":
                    temp_file_path = f"temp_{file.filename}"
                    with open(temp_file_path, "wb") as temp_file:
                        temp_file.write(file.file.read())
                    extracted_text = read_doc(temp_file_path)
                    os.remove(temp_file_path)
                else:
                    raise ValueError("Unsupported Word file format.")
                response_data[file.filename] = {"content": extracted_text}
            elif file.content_type == "text/plain" or file_extension == "txt":
                extracted_text = read_txt(io.BytesIO(file.file.read()))
                response_data[file.filename] = {"content": extracted_text}
            else:
                response_data[file.filename] = {
                    "error": "Unsupported file type. Supported formats: .pdf, .docx, .doc, .txt, .zip, .rar"
                }

        except Exception as e:
            response_data[file.filename] = {
                "error": str(e)
            }
    conversation_resume = get_conversation(TEMPLATES["resume"])
    conversation_score = get_conversation(TEMPLATES["score"])
    for filename in response_data:
        print(f"{filename}") #: {response_data[filename]["content"]}

        resume_response = conversation_resume.invoke({"resume_text": response_data[filename]["content"]})
        time.sleep(3)
        # print(resume_response.content)
        response_data[filename]["key_feature"] = resume_response.content
        # response_data[filename] = {"key_aspect": resume_response.content}
        score_response = conversation_score.invoke({
            "resume_text": resume_response.content,
            "job_description": processed_jd
        })
        time.sleep(3)
        response_data[filename]["score"] = score_response.content
        print(score_response.content)

    return response_data







# from fastapi import FastAPI, File, UploadFile
# from fastapi.responses import JSONResponse
# from PyPDF2 import PdfReader
# from docx import Document
# import win32com.client as win32
# import os
# import io

# app = FastAPI()

# def read_pdf(file: UploadFile):
#     """Extract text from a PDF file."""
#     contents = file.file.read()
#     pdf_reader = PdfReader(io.BytesIO(contents))
#     extracted_text = ""
#     for page in pdf_reader.pages:
#         extracted_text += page.extract_text()
#     return extracted_text.strip()

# def read_docx(file: UploadFile):
#     """Extract text from a DOCX file."""
#     contents = file.file.read()
#     document = Document(io.BytesIO(contents))
#     extracted_text = ""
#     for paragraph in document.paragraphs:
#         extracted_text += paragraph.text + "\n"
#     return extracted_text.strip()

# def read_doc(file: UploadFile):
#     """
#     Extract text from a DOC file using COM automation (Windows only).
#     """
#     try:
#         # Save the uploaded file temporarily
#         temp_doc_path = file.filename
#         with open(temp_doc_path, "wb") as temp_file:
#             temp_file.write(file.file.read())

#         # Use COM automation to extract text
#         word = win32.Dispatch("Word.Application")
#         word.Visible = False
#         doc = word.Documents.Open(os.path.abspath(temp_doc_path))
#         extracted_text = doc.Content.Text
#         doc.Close(False)
#         word.Quit()

#         # Clean up the temporary file
#         os.remove(temp_doc_path)

#         return extracted_text.strip()
#     except Exception as e:
#         return f"Error processing DOC file: {str(e)}"

# def read_txt(file: UploadFile):
#     """Extract text from a plain text file."""
#     try:
#         contents = file.file.read()
#         return contents.decode("utf-8").strip()
#     except Exception as e:
#         return f"Error processing TXT file: {str(e)}"

# @app.post("/upload-files/")
# async def upload_files(files: list[UploadFile] = File(...)):
#     response_data = []

#     for file in files:
#         try:
#             if file.content_type == "application/pdf":
#                 extracted_text = read_pdf(file)
#             elif file.content_type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"]:
#                 if file.filename.endswith(".docx"):
#                     extracted_text = read_docx(file)
#                 elif file.filename.endswith(".doc"):
#                     extracted_text = read_doc(file)
#                 else:
#                     raise ValueError("Unsupported Word file format.")
#             elif file.content_type == "text/plain":
#                 extracted_text = read_txt(file)
#             else:
#                 return JSONResponse(
#                     content={"error": f"Unsupported file type for {file.filename}. Supported formats are .pdf, .docx, .doc, and .txt."},
#                     status_code=400,
#                 )

#             response_data.append({
#                 "filename": file.filename,
#                 "content": extracted_text
#             })

#         except Exception as e:
#             response_data.append({
#                 "filename": file.filename,
#                 "error": str(e)
#             })

#     return response_data
