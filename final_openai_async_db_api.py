from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, FileResponse
from PyPDF2 import PdfReader
from docx import Document
import win32com.client
import os
import openai
import io
import zipfile
import pandas as pd
import time
from dotenv import load_dotenv
from langchain_core.prompts import PromptTemplate
from datetime import datetime
import asyncio
import re
import psycopg2
import shutil
import utils

# Load API Key
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

openai.api_key = API_KEY

hostname = 'localhost'
username = 'postgres'
password = '123456'
database = 'ResumeDB'
port_id = 5432

conn = None
cur = None

app = FastAPI()

def retrieve_resume_blob(unique_id):
    conn = None
    try:
        # Connect to the PostgreSQL server
        conn = psycopg2.connect(
            host=hostname,
            dbname=database,
            user=username,
            password=password,
            port=port_id
        )

        # Create a cursor
        cur = conn.cursor()

        # SQL query to retrieve the BLOB data
        cur.execute(
            "SELECT unique_id, resume_name, blob_data FROM resume_table WHERE unique_id = %s", 
            (unique_id,)
        )

        # Fetch the result
        result = cur.fetchone()

        if result:
            unique_id, resume_name, blob_data = result
            
            # Define the output directory (you can modify this path as needed)
            output_dir = r"extracted_files"
            
            # Create the directory if it doesn't exist
            import os
            os.makedirs(output_dir, exist_ok=True)

            # Full path for the output file
            output_path = os.path.join(output_dir, f"{unique_id}_{resume_name}")

            # Write the BLOB data to a file
            with open(output_path, 'wb') as file:
                file.write(blob_data)

            print(f"Resume retrieved and saved to: {output_path}")
            return output_path
        else:
            print(f"No resume found with unique_id: {unique_id}")
            return None

    except (Exception, psycopg2.DatabaseError) as error:
        print(f"Error retrieving resume: {error}")
    finally:
        if cur is not None:
            cur.close()
            print('Cursor closed.')

        if conn is not None:
            conn.close()
            print('Database connection closed.')

    return output_path


conversation_resume = utils.get_conversation_openai(utils.TEMPLATES["resume"])
conversation_score = utils.get_conversation_openai(utils.TEMPLATES["score"])

async def run_in_executor(func, *args, **kwargs):
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, func, *args, **kwargs)

async def async_key_aspect_extractor(filename, data):
    try:
        print(f"Extracting key aspects for: {filename} - START")
        # Assuming conversation_resume is synchronous
        result = await run_in_executor(conversation_resume, {"resume_text": data["content"]})
        return filename, result
    except Exception as e:
        print(f"Error in key aspect extraction for {filename}: {e}")
        return filename, None

async def async_resume_scorer(filename, key_aspect, job_description):
    try:
        print(f"Scoring resume: {filename} - START")
        # Assuming conversation_score is synchronous
        result = await run_in_executor(conversation_score, {
            "resume_text": key_aspect,
            "job_description": job_description
        })
        return filename, result
    except Exception as e:
        print(f"Error in scoring for {filename}: {e}")
        return filename, None

async def process_resumes_async(response_data, job_description):
    # Create async tasks for key aspect extraction
    key_aspect_tasks = [
        asyncio.create_task(async_key_aspect_extractor(filename, data)) 
        for filename, data in response_data.items()
    ]
    
    # Wait for all key aspect extraction tasks to complete
    key_aspects = await asyncio.gather(*key_aspect_tasks, return_exceptions=True)
    key_aspects_dict = {filename: result for filename, result in key_aspects if result is not None}
    
    # Create async tasks for scoring
    scoring_tasks = [
        asyncio.create_task(async_resume_scorer(filename, key_aspects_dict.get(filename, ""), job_description)) 
        for filename in response_data.keys()
    ]
    
    # Wait for all scoring tasks to complete concurrently
    scores = await asyncio.gather(*scoring_tasks, return_exceptions=True)
    scores_dict = {filename: result for filename, result in scores if result is not None}
    
    # Update response_data with results
    for filename in response_data.keys():
        response_data[filename]['key_feature'] = utils.clean_text(key_aspects_dict.get(filename, ""))
        response_data[filename]['score'] = utils.clean_text(scores_dict.get(filename, ""))
    
    return response_data

@app.post("/upload-files/")
async def upload_files(job_description: str, files: list[UploadFile] = File(...)):
    response_data = {}

    conversation_jd = utils.get_conversation_openai(utils.TEMPLATES["job_description"])
    jd_response = conversation_jd({"job_description_text": job_description})
    processed_jd = jd_response
    print("Processing the Job Description...\n")

    # Create extracted_files directory if it doesn't exist
    extract_path = "extracted_files"
    
    # Clean the extracted_files directory before processing
    if os.path.exists(extract_path):
        shutil.rmtree(extract_path)
    
    os.makedirs(extract_path, exist_ok=True)

    try:
        conn = psycopg2.connect(
        host=hostname,
        user=username,
        password=password,
        dbname=database,
        port=port_id
        )

        cur = conn.cursor()

        cur.execute("""
                CREATE TABLE IF NOT EXISTS resume_table (
                    unique_id NUMERIC PRIMARY KEY,
                    resume_name VARCHAR(100) ,
                    resume_content TEXT ,
                    resume_key_aspect TEXT ,
                    score INTEGER ,
                    blob_data BYTEA 
                )
            """)
        conn.commit()

    except (Exception, psycopg2.Error) as error:
        print("Error while connecting to PostgreSQL", error)

    for file in files:
        try:
            # Fallback to file extension if MIME type is not reliable
            file_extension = file.filename.split(".")[-1].lower()

            # Generate a unique file name using a timestamp
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
            file_name = file.filename
            unique_filename = f"{timestamp}_{file_name}"
            file_path = os.path.join(extract_path, unique_filename)

            if file.content_type == "application/zip" or file_extension == "zip":
                zip_data = await file.read()
                zip_file = io.BytesIO(zip_data)
                extracted_files = []
                with zipfile.ZipFile(zip_file, 'r') as z:
                    # Extract files
                    file_name_list = z.namelist()
                    for original_file_name in file_name_list:
                        # Generate a unique file name for each file in the ZIP
                        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
                        unique_file_name = f"{timestamp}_{original_file_name}"
                        z.extract(original_file_name, extract_path)
                        # Rename the extracted file to its unique name
                        os.rename(os.path.join(extract_path, original_file_name),
                                  os.path.join(extract_path, unique_file_name))
                        extracted_files.append(unique_file_name)

                pdf_contents = {}
                i = 0
                for file_name in extracted_files:
                    original_name = file_name_list[i]
                    i += 1
                    print(f"Reading file: {original_name}")
                    file_path = os.path.join(extract_path, file_name)
                    unique_id = re.match(r'^\d+', file_name).group()
                    resume_name = original_name
                    resume_content = None
                    blob_data = None

                    if file_name.endswith(".pdf"):
                        with open(file_path, "rb") as pdf_file:
                            # First read the entire file content as binary
                            data = pdf_file.read()
                            # Create blob data from the binary content
                            blob_data = psycopg2.Binary(data)
                            # Reset file pointer to start
                            pdf_file.seek(0)
                            # Now read for text extraction
                            resume_content = utils.read_pdf(pdf_file)
                            response_data[original_name] = {"content": resume_content, "file_path": unique_id}
                    
                    # Process TXT files
                    elif file_name.endswith(".txt"):
                        with open(file_path, "rb") as txt_file:
                            # First read the entire file content as binary
                            data = txt_file.read()
                            # Create blob data from the binary content
                            blob_data = psycopg2.Binary(data)
                            # Reset file pointer to start
                            txt_file.seek(0)
                            # Now read for text extraction
                            resume_content = utils.read_txt(txt_file)
                            response_data[original_name] = {"content": resume_content, "file_path": unique_id}

                    elif file_name.endswith(".docx"):
                        try:
                            with open(file_path, "rb") as docx_file:
                                # First read the entire file content as binary
                                data = docx_file.read()
                                # Create blob data from the binary content
                                blob_data = psycopg2.Binary(data)
                                # Reset file pointer to start
                                docx_file.seek(0)
                                # Now read for text extraction
                                resume_content = utils.read_docx(docx_file)
                                response_data[original_name] = {"content": resume_content, "file_path": unique_id}
                        except Exception as e:
                            response_data[original_name] = {"content": f"Error reading DOCX file: {str(e)}", "file_path": unique_id}

                    # Process DOC files
                    elif file_name.endswith(".doc"):
                        resume_content, blob_data = utils.read_doc(file_path)
                        response_data[original_name] = {"content": resume_content, "file_path": unique_id}
                
                    if resume_content is not None and blob_data is not None:
                        try:
                            # Insert the data into the database
                            cur.execute("""
                                INSERT INTO resume_table (unique_id, resume_name, resume_content, resume_key_aspect, score, blob_data)
                                VALUES (%s, %s, %s, %s, %s, %s)
                            """, (unique_id, resume_name, resume_content, None, None, blob_data))
                            conn.commit()
                            print(f"Successfully stored {resume_name} in database")
                        except Exception as e:
                            print(f"Error storing {resume_name} in database: {str(e)}")
                            conn.rollback()
                    else:
                        print(f"Skipping {resume_name} - No content or blob data available")

                    # Clean up the extracted file
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            print(f"Error removing temporary file {file_path}: {str(e)}")

            else:
                unique_id = timestamp
                resume_name = file_name
                resume_content = None
                blob_data = None

                # Save individual file to extracted_files directory
                file_content = await file.read()
                
                with open(file_path, "wb") as f:
                    f.write(file_content)

                print(f"Reading file: {file_name}")
                # Process based on file type
                if file_extension == "pdf":
                    with open(file_path, "rb") as pdf_file:
                        # First read the entire file content as binary
                        data = pdf_file.read()
                        # Create blob data from the binary content
                        blob_data = psycopg2.Binary(data)
                        # Reset file pointer to start
                        pdf_file.seek(0)
                        # Now read for text extraction
                        resume_content = utils.read_pdf(pdf_file)
                        response_data[file_name] = {"content": resume_content}
                
                elif file_extension == "txt":
                    with open(file_path, "rb") as txt_file:
                        # First read the entire file content as binary
                        data = txt_file.read()
                        # Create blob data from the binary content
                        blob_data = psycopg2.Binary(data)
                        # Reset file pointer to start
                        txt_file.seek(0)
                        # Now read for text extraction
                        resume_content = utils.read_txt(txt_file)
                        response_data[file_name] = {"content": resume_content}
                        
                
                elif file_extension == "docx":
                    try:
                        with open(file_path, "rb") as docx_file:
                            # First read the entire file content as binary
                            data = docx_file.read()
                            # Create blob data from the binary content
                            blob_data = psycopg2.Binary(data)
                            # Reset file pointer to start
                            docx_file.seek(0)
                            # Now read for text extraction
                            resume_content = utils.read_docx(docx_file)
                            response_data[file_name] = {"content": resume_content}
                            
                    except Exception as e:
                        response_data[file_name] = {"content": str(e)}
                
                elif file_extension == "doc":
                    resume_content, blob_data = utils.read_doc(file_path)
                    response_data[file_name] = {"content": resume_content}
                
                else:
                    # response_data[file_name] = {
                    #     "error": "Unsupported file type. Supported formats: .pdf, .docx, .doc, .txt, .zip, .rar"
                    # }
                    pass
                # Add file path to the response data
                response_data[file_name]["file_path"] = unique_id
            
                # SQL query to insert data into the database. 
                cur.execute( 
                        "INSERT INTO resume_table(unique_id,resume_name,resume_content,blob_data) "
                        "VALUES(%s,%s,%s,%s)", (unique_id, resume_name, resume_content, blob_data)
                        )

                conn.commit()

                # Clean up the extracted file
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        print(f"Error removing temporary file {file_path}: {str(e)}")

        except Exception as e:
            response_data[file.filename] = {
                "error": str(e)
            }
        
    print("\n")       
    response_data = await process_resumes_async(response_data, processed_jd)

    for key, value in response_data.items():
        resume_key_aspect = value["key_feature"]
        score = value["score"]
        unique_id = value["file_path"]

        cur.execute(
                """
                UPDATE resume_table 
                SET resume_key_aspect = %s, 
                    score = %s 
                WHERE unique_id = %s
                """, 
                (resume_key_aspect, score, unique_id)
            )

            # Commit the changes
        conn.commit()

    resume_df = pd.DataFrame(columns=['Resume Name', 'Score'])
    i = 0
    for key, value in response_data.items():
        resume_df.loc[i, "Resume Name"] = key
        resume_df.loc[i, "Score"] = value["score"]
        i += 1

    resume_df.sort_values(by='Score', ascending=False, inplace=True)
    file_path = os.path.join("extracted_files", f'R_Resume_Scorecard.xlsx')

        # Save the scorecard to an Excel file
    resume_df.to_excel(file_path, index=False)

    if cur is not None:
        cur.close()
        print('Cursor closed.')

    if conn is not None:
        conn.close()
        print('Database connection closed.')

    return response_data


@app.post("/download-resume/{file_path}")
def download_file(file_path: str):
    """Endpoint to download a file by its name."""
    # file_path = os.path.join("extracted_files", file_name)  # Adjust the path as needed
    file_path = retrieve_resume_blob(file_path)
    if os.path.exists(file_path):
        print("Dowloading the file...") 
        return FileResponse(file_path, media_type='application/octet-stream', filename = file_path.split('_', 2)[-1])
    else:
        return JSONResponse(content={"message": "File not found."}, status_code=404)
        

@app.post("/download-scorecard")
def download_file():
    """Endpoint to download the excel file containing resume scores."""
    # file_path = os.path.join("extracted_files", file_name)  # Adjust the path as needed
    file_path = os.path.join("extracted_files", f'R_Resume_Scorecard.xlsx')
    if os.path.exists(file_path):
        print("Dowloading the file...") 
        return FileResponse(file_path, media_type='application/octet-stream', filename = file_path.split('_', 2)[-1])
    else:
        return JSONResponse(content={"message": "File not found."}, status_code=404)
