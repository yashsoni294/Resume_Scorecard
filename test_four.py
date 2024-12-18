from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from datetime import datetime
import os
import asyncio
import uuid
import psycopg2
from psycopg2 import pool
import zipfile
import io
import re
import logging
from typing import List
import aiofiles

# Configure logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Database connection pool configuration
class DatabaseConnectionPool:
    _instance = None
    
    def __new__(cls):
        if not cls._instance:
            try:
                cls._instance = super().__new__(cls)
                cls._instance.connection_pool = psycopg2.pool.SimpleConnectionPool(
                    1,  # min connections
                    20,  # max connections
                    host='localhost',
                    user='postgres',
                    password='123456',
                    database='ResumeDB',
                    port=5432
                )
                logger.info("Database connection pool created successfully")
            except Exception as e:
                logger.error(f"Error creating database connection pool: {e}")
                raise
        return cls._instance
    
    def get_connection(self):
        try:
            return self.connection_pool.getconn()
        except Exception as e:
            logger.error(f"Error getting database connection: {e}")
            raise
    
    def release_connection(self, conn):
        try:
            self.connection_pool.putconn(conn)
        except Exception as e:
            logger.error(f"Error releasing database connection: {e}")

# Async file handling utility
class AsyncFileHandler:
    @staticmethod
    async def save_uploaded_file(file: UploadFile, base_path: str) -> str:
        """
        Asynchronously save an uploaded file with a unique name
        
        :param file: Uploaded file object
        :param base_path: Base directory to save the file
        :return: Full path of the saved file
        """
        try:
            # Ensure base path exists
            os.makedirs(base_path, exist_ok=True)
            
            # Generate a unique filename
            unique_filename = f"{uuid.uuid4()}_{file.filename}"
            full_path = os.path.join(base_path, unique_filename)
            
            # Asynchronously write file
            async with aiofiles.open(full_path, 'wb') as out_file:
                content = await file.read()
                await out_file.write(content)
            
            return full_path
        except Exception as e:
            logger.error(f"Error saving file {file.filename}: {e}")
            raise HTTPException(status_code=500, detail=f"File save failed: {e}")

# Main FastAPI application
app = FastAPI(title="Multi-User Resume Processing")

# CORS configuration
origins = [
    "http://localhost:3000",
    "https://your-frontend-domain.com"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload-files/")
async def upload_files(
    job_description: str, 
    files: List[UploadFile] = File(...)
):
    """
    Enhanced multi-user file upload endpoint with robust concurrency handling
    
    :param job_description: Job description for resume scoring
    :param files: List of uploaded files
    :return: Processing results
    """
    # Generate a unique session ID for this upload batch
    session_id = str(uuid.uuid4())
    base_upload_path = os.path.join("uploads", session_id)
    
    try:
        # Validate inputs
        if not files:
            raise HTTPException(status_code=400, detail="No files uploaded")
        
        # Concurrent file saving
        saved_files = await asyncio.gather(
            *[AsyncFileHandler.save_uploaded_file(file, base_upload_path) for file in files]
        )
        
        # Process files concurrently
        processing_results = await process_files_concurrently(saved_files, job_description)
        
        return JSONResponse(
            status_code=200, 
            content={
                "session_id": session_id,
                "processed_files": processing_results
            }
        )
    
    except Exception as e:
        logger.error(f"Upload process failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

async def process_files_concurrently(file_paths: List[str], job_description: str):
    """
    Process multiple files concurrently with proper error handling
    
    :param file_paths: List of file paths to process
    :param job_description: Job description for scoring
    :return: Processing results for each file
    """
    # Create a semaphore to limit concurrent processing
    semaphore = asyncio.Semaphore(5)  # Limit to 5 concurrent file processing tasks
    
    async def process_single_file(file_path):
        async with semaphore:
            try:
                # Implement your file processing logic here
                # This is a placeholder - replace with actual processing
                return {
                    "file_path": file_path,
                    "status": "processed",
                    "details": "File processed successfully"
                }
            except Exception as e:
                logger.error(f"Error processing {file_path}: {e}")
                return {
                    "file_path": file_path,
                    "status": "failed",
                    "error": str(e)
                }
    
    # Process files concurrently
    results = await asyncio.gather(
        *[process_single_file(file_path) for file_path in file_paths]
    )
    
    return results

# Startup and shutdown events for resource management
@app.on_event("startup")
async def startup_event():
    # Initialize database connection pool
    try:
        DatabaseConnectionPool()
        logger.info("Application startup completed")
    except Exception as e:
        logger.error(f"Startup failed: {e}")

@app.on_event("shutdown")
async def shutdown_event():
    # Close database connection pool if needed
    logger.info("Application is shutting down")