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
import boto3
from botocore.exceptions import NoCredentialsError, ClientError
import os

# Load API Key
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

# AWS Credentials from environment variables
AWS_ACCESS_KEY_ID = os.getenv('AWS_ACCESS_KEY_ID')
AWS_SECRET_ACCESS_KEY = os.getenv('AWS_SECRET_ACCESS_KEY')
AWS_REGION = os.getenv('AWS_REGION', 'ap-south-1')  # Default region if not specified

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

def create_s3_client():
    """
    Create and return an S3 client with configured credentials
    
    :return: Boto3 S3 client
    """
    return boto3.client(
        's3',
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=AWS_REGION
    )

def upload_to_s3(local_file_path, bucket_name='yash-soni-db', s3_folder='resume_files/'):
    """
    Upload a file to an S3 bucket
    
    :param local_file_path: Path to the local file to upload
    :param bucket_name: Name of the S3 bucket
    :param s3_folder: Folder path within the bucket (include trailing '/')
    :return: True if file was uploaded, else False
    """
    # Create an S3 client
    s3 = create_s3_client()
    
    try:
        # Get the filename from the path
        filename = os.path.basename(local_file_path)
        
        # Construct the full S3 key (path) 
        s3_key = os.path.join(s3_folder, filename)
        
        # Upload the file
        s3.upload_file(local_file_path, bucket_name, s3_key)
        print(f"Successfully uploaded {filename} to {bucket_name}/{s3_key}")
        return True
    
    except FileNotFoundError:
        print(f"The file {local_file_path} was not found")
        return False
    
    except NoCredentialsError:
        print("Credentials not available")
        return False
    
    except ClientError as e:
        print(f"An error occurred: {e}")
        return False

def download_from_s3(filename, bucket_name='yash-soni-db', s3_folder='resume_files/', 
                     local_dir='extracted_files/'):
    """
    Download a specific file from S3 bucket
    
    :param filename: Name of the file to download
    :param bucket_name: Name of the S3 bucket
    :param s3_folder: Folder path within the bucket (include trailing '/')
    :param local_dir: Local directory to save the downloaded file
    :return: Path to the downloaded file or None if download fails
    """
    # Create an S3 client
    s3 = create_s3_client()
    
    # Create local directory if it doesn't exist
    os.makedirs(local_dir, exist_ok=True)
    
    try:
        # Construct the full S3 key (path)
        s3_key = os.path.join(s3_folder, filename)
        
        # Local file path
        local_file_path = os.path.join(local_dir, filename)
        
        # Download the file
        s3.download_file(bucket_name, s3_key, local_file_path)
        
        print(f"Successfully downloaded {filename} to {local_file_path}")
        return local_file_path
    
    except ClientError as e:
        print(f"Error downloading {filename}: {e}")
        return None


def upload_resume_file(filename, directory_path='extracted_files'):
    """
    Upload a specific file from a directory to S3
    
    :param filename: Name of the file to upload
    :param directory_path: Path to the directory containing the file
    :return: True if file was uploaded successfully, False otherwise
    """
    # Construct full file path
    file_path = os.path.join(directory_path, filename)
    
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File {filename} not found in {directory_path}")
        return False
    
    # Check if it's a file (not a directory)
    if not os.path.isfile(file_path):
        print(f"{filename} is not a file")
        return False
    
    # Attempt to upload the file
    return upload_to_s3(file_path)


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
                    score INTEGER
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
            time.sleep(0.001)
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
                        time.sleep(0.001)
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
                    

                    if file_name.endswith(".pdf"):
                        with open(file_path, "rb") as pdf_file:
                            # Now read for text extraction
                            resume_content = utils.read_pdf(pdf_file)
                            response_data[original_name] = {"content": resume_content, "file_path": file_name}
                    
                    # Process TXT files
                    elif file_name.endswith(".txt"):
                        with open(file_path, "rb") as txt_file:
                            # Now read for text extraction
                            resume_content = utils.read_txt(txt_file)
                            response_data[original_name] = {"content": resume_content, "file_path": file_name}

                    elif file_name.endswith(".docx"):
                        try:
                            with open(file_path, "rb") as docx_file:
                                # Now read for text extraction
                                resume_content = utils.read_docx(docx_file)
                                response_data[original_name] = {"content": resume_content, "file_path": file_name}
                        except Exception as e:
                            response_data[original_name] = {"content": f"Error reading DOCX file: {str(e)}", "file_path": file_name}

                    # Process DOC files
                    elif file_name.endswith(".doc"):
                        resume_content, _ = utils.read_doc(file_path)
                        response_data[original_name] = {"content": resume_content, "file_path": file_name}
                
                    if resume_content is not None:
                        try:
                            # Insert the data into the database
                            cur.execute("""
                                INSERT INTO resume_table (unique_id, resume_name, resume_content, resume_key_aspect, score)
                                VALUES (%s, %s, %s, %s, %s)
                            """, (unique_id, resume_name, resume_content, None, None))
                            conn.commit()
                            print(f"Successfully stored {resume_name} in database")
                        except Exception as e:
                            print(f"Error storing {resume_name} in database: {str(e)}")
                            conn.rollback()
                    else:
                        print(f"Skipping {resume_name} - No content or blob data available")

                    upload_resume_file(filename=file_name, directory_path='extracted_files')
                    print("Uploaded to S3 Bucket")


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


                # Save individual file to extracted_files directory
                file_content = await file.read()
                
                with open(file_path, "wb") as f:
                    f.write(file_content)

                print(f"Reading file: {file_name}")
                # Process based on file type
                if file_extension == "pdf":
                    with open(file_path, "rb") as pdf_file:
                        # Now read for text extraction
                        resume_content = utils.read_pdf(pdf_file)
                        response_data[file_name] = {"content": resume_content}
                
                elif file_extension == "txt":
                    with open(file_path, "rb") as txt_file:
                        # Now read for text extraction
                        resume_content = utils.read_txt(txt_file)
                        response_data[file_name] = {"content": resume_content}
                        
                
                elif file_extension == "docx":
                    try:
                        with open(file_path, "rb") as docx_file:
                            # Now read for text extraction
                            resume_content = utils.read_docx(docx_file)
                            response_data[file_name] = {"content": resume_content}
                            
                    except Exception as e:
                        response_data[file_name] = {"content": str(e)}
                
                elif file_extension == "doc":
                    resume_content, _ = utils.read_doc(file_path)
                    response_data[file_name] = {"content": resume_content}
                
                else:
                    # response_data[file_name] = {
                    #     "error": "Unsupported file type. Supported formats: .pdf, .d ocx, .doc, .txt, .zip, .rar"
                    # }
                    pass
                # Add file path to the response data
                response_data[file_name]["file_path"] = f"{unique_id}_{resume_name}"

                upload_resume_file(filename = f"{unique_id}_{resume_name}", directory_path='extracted_files')
                print("Uploaded to S3 Bucket")
            
                # SQL query to insert data into the database. 
                cur.execute( 
                        "INSERT INTO resume_table(unique_id,resume_name,resume_content) "
                        "VALUES(%s,%s,%s)", (unique_id, resume_name, resume_content)
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
        unique_id = re.match(r"^\d{20}", unique_id).group()


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
    # file_path = retrieve_resume_blob(file_path)
    file_path = download_from_s3(file_path)
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
