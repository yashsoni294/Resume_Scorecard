import openai
from langchain.prompts import PromptTemplate
import os
from dotenv import load_dotenv
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
from langchain_core.prompts import PromptTemplate
from datetime import datetime

# Load API Key
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

openai.api_key = API_KEY


# Constants
TEMPLATES = {
    "job_description": """
        
        The text below is a job description:

        {job_description_text}

        Your task is to analyze the job description and extract key aspects that will help evaluate a candidate's suitability. Organize the extracted information into the following categories to support scoring. Also Avoid including unnecessary details or assumptions on your own :

        1. Candidate Profile :-
            Job-Related Keywords:
            Extract the most relevant keywords and phrases used in the job description, such as specific skills, technologies, or qualifications, that highlight the core expectations for the role.

            Relevant Past Roles and Responsibilities:
            Identify the types of roles and responsibilities that align closely with this position, based on the job description.

            Actionable Responsibilities:
            List clear, action-oriented tasks or expectations (e.g., "Develop and manage X," "Collaborate on Y") that indicate measurable outcomes for success in this role.

        2. Experience Section :-
            Years of Experience:
            Specify the minimum and preferred years of experience required or desired for the role.

            Technical Skills:
            Extract a comprehensive list of both mandatory and preferred technical skills, including tools, technologies, programming languages, or domain-specific expertise mentioned in the job description.

            Soft Skills and Teamwork:
            Identify soft skills and teamwork-related requirements (e.g., leadership, communication, collaboration), with examples if provided in the description.

        3. Educational Qualifications and Certifications :- 
            Minimum Educational Qualifications:
            Note the required or preferred educational qualifications for this role.

            Relevant Certifications/Training Programs:
            List certifications, licenses, or training programs that are directly or partially relevant to the job description.

            Output:
            Provide the extracted information in a concise, bullet-pointed format, ensuring relevance to the scoring criteria. Avoid including unnecessary details or assumptions on your own. This structured output will guide the scoring process.        
            """,

    "resume":""" 
            The text below is a resume:

            {resume_text}

            Your task is to extract information from the resume by focusing on the following areas. Ensure the extracted information is clear, structured, and concise also Ensure the output focuses only on the resume content without making assumptions or adding extra information from your own :

            ### 1. **Candidate Profile**
            - **Keywords Identified:** List relevant keywords found in the resume that reflect the candidate skills, roles, and expertise.  
            - **Past Roles Summary:** Provide an overview of the candidate's past roles and responsibilities, emphasizing their clarity and detail. Highlight the use of action words (e.g., "Developed," "Managed") and measurable outcomes.  
            - **Clarity of Responsibilities:** Note how clearly responsibilities are described, focusing on structured descriptions and measurable achievements.

            ### 2. **Experience Section**
            - **Years of Experience:** Indicate the total years of experience and specify industries or domains the candidate has worked in.  
            - **Technical Skills:** Identify technical skills explicitly mentioned in the resume, along with any supporting examples or certifications.  
            - **Communication and Teamwork Skills:** Highlight references to teamwork and communication skills, particularly examples such as leadership roles, collaboration efforts, or achievements in group settings.

            ### 3. **Educational Qualifications and Certifications**
            - **Educational Background:** Provide the highest qualification achieved and any notable academic achievements.  
            - **Certifications and Training:** List certifications or training programs mentioned in the resume, along with their relevance to enhancing the candidate expertise.

            Provide a structured summary of the extracted information in bullet points or short sentences for each section.

            """,
    "score": """
        Your task is to evaluate how well the provided resume aligns with the given job description by assessing three key factors: 1. Candidate Profile 2. Experience section 3. Educational Qualifications and Certifications. Based on this evaluation, provide a final score between 0 and 100, representing the overall suitability of the candidate for the job.

        ### Inputs:
        - **Resume Text:**  

        {resume_text}

        - **Job Description Text:**  

        {job_description}

        ### Scoring Guidelines:
        Give marks to the resume against the job description across the following criterias :-

        1. Candidate Profile (Max 15 Marks) :-

            •	Job-Related Keywords (Max 5 Marks):
                o	5 Points: IF Resume includes highly relevant keywords from the Job Description, indicating alignment with job requirements.
                o	3 Points: IF Resume Includes some relevant keywords but lacks critical ones from the Job Description.
                o	1 Point: IF Few or no relevant keywords used.
            •	Relevance of Past Roles to Job Description (Max 5 Marks):
                o	5 Points: IF Past roles and responsibilities align strongly with Job Description requirements.
                o	3 Points: IF Moderate alignment, some responsibilities match the Job Description.
                o	1 Point: IF Weak or no relevance to the Job Description.
            •	Clarity of Responsibilities (Max 5 Marks):
                o	5 Points: IF Responsibilities are clearly defined using action words (e.g., "Developed," "Managed") and measurable outcomes.
                o	3 Points: IF Responsibilities are described but lack action words or outcomes.
                o	1 Point: IF Responsibilities are vague or generic.
        
        2. Experience section (Max 65) :- 

            • Years of Experience (Max 15 Marks):
                o	15 Points: IF Experience matches or exceeds the Job Description requirements.
                o	10 Points: IF Slightly below required years but relevant experience.
                o	5 Points: IF Limited relevance or inadequate years of experience.
            • Matching Technical Skills (Max 40 Marks):
                o	40 Points: IF All technical skills mentioned in the Job Description are evident, with examples or certifications provided.
                o	30 Points: IF Most technical skills are evident but lack depth or examples.
                o	20 Points: IF Some technical skills match; others are missing.
                o	10 Points or Below: IF Minimal or no alignment with required technical skills.
            • Communication and Teamwork (Max 10 Marks):
                o	10 Points: IF Resume highlights soft skills with examples (e.g., "Led a team of 5," "Facilitated cross-functional communication").
                o	6 Points: IF Mentions soft skills but lacks examples.
                o	3 Points: IF Minimal mention of soft skills.

        3. Educational Qualifications and Certifications (Max 20) :- 
                
            • Meets Minimum Educational Qualifications (Max 15 Marks):
                o	15 Points: IF Meets or exceeds educational qualifications specified in the Job Description.
                o	10 Points: IF Meets basic qualifications but lacks advanced or desired qualifications.
                o	5 Points: IF Does not fully meet educational qualifications.
            • Additional Certifications/Training Programs (Max 5 Marks):
                o	5 Points: IF Certifications/training directly related to Job Description (e.g., HR certifications, technical tools training).
                o	3 Points: IF Certifications or training partially relevant to the Job Description.
                o	1 Point: IF No additional certifications.

        Add marks of all Three section (Candidate Profile, Experience section, Educational Qualifications and Certifications) Round the result to the nearest whole number.

        ### Output:
        Provide only the final average score as a single number (0-100) with no additional text or explanation.
    """
    ,
}


# Initialize DataFrame
resume_df = pd.DataFrame(columns=['resume_file_name', 'resume_file_text', 'resume_key_aspect', 'resume_score'])


def get_conversation_openai(template, model="gpt-4o-mini", temperature=0.1, max_tokens=None):
 
    # Initialize the OpenAI model
    def call_openai_model(inputs):
        """
        Invokes the OpenAI model using the provided inputs.
        """
        prompt = PromptTemplate.from_template(template).format(**inputs)
        response = openai.ChatCompletion.create(
            model=model,
            messages=[{"role": "system", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens
        )
        return response["choices"][0]["message"]["content"]
    
    return call_openai_model

def select_folder():

    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder").replace('/', '\\')

def read_pdf(file_path):

    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
    return text

def read_txt(file_path):

    text = ""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
    except Exception as e:
        print(f"Error reading TXT file {file_path}: {e}")
    return text

def read_docx(file_path):

    try:
        doc = docx.Document(file_path)
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return ""

def read_doc(file_path):

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        return text
    except Exception as e:
        print(f"Error reading DOC {file_path}: {e}")
        return ""



def extract_text_from_files(folder_path):

    data = []
    for filename in set(os.listdir(folder_path)):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.pdf'):
            text = read_pdf(file_path)
        elif filename.endswith('.docx'):
            text = read_docx(file_path)
        elif filename.endswith('.doc'):
            text = read_doc(file_path)
        elif filename.endswith('.txt'):
            text = read_txt(file_path)
        else:
            continue
        data.append({'resume_file_name': filename, 'resume_file_text': text})
    return pd.DataFrame(data)

def process_resumes(resume_df, job_description):
  

    conversation_jd = get_conversation_openai(TEMPLATES["job_description"])
    # conversation_jd = get_conversation(TEMPLATES["job_description"])
    # jd_response = conversation_jd.invoke({"job_description_text": job_description})
    jd_response = conversation_jd({"job_description_text": job_description})
    processed_jd = jd_response

    # conversation_resume = get_conversation(TEMPLATES["resume"])
    conversation_resume = get_conversation_openai(TEMPLATES["resume"])
    # conversation_score = get_conversation(TEMPLATES["score"])
    conversation_score = get_conversation_openai(TEMPLATES["score"])

    for i in range(len(resume_df)):
        resume_text = resume_df.loc[i, "resume_file_text"]
        
        # Extract key aspects
        # resume_response = conversation_resume.invoke({"resume_text": resume_text})
        resume_response = conversation_resume({"resume_text": resume_text})
        resume_df.loc[i, "resume_key_aspect"] = resume_response
        
        # time.sleep(3)  # Avoid API rate limits
        
        # Score resume
        score_response = conversation_score({
            "resume_text": resume_text,
            "job_description": processed_jd
        })
        resume_df.loc[i, "resume_score"] = score_response

        print(f"{i+1}. Working on - ",resume_df["resume_file_name"][i])

        # time.sleep(3)  # Avoid API rate limits
    return resume_df

def save_results(resume_df):

    resume_df.sort_values(by='resume_score', ascending=False, inplace=True)
    # resume_df.to_pickle("resume_df.pkl")

    scorecard = resume_df[["resume_file_name", "resume_score"]]

    # Get the user's Downloads folder
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    # Get the current date and time
    # now = datetime.now()

    # Format the time
    current_time = datetime.now().strftime("%d-%b-%Y_%I-%M-%S_%p")

    # # Define the full file path
    file_path = os.path.join(downloads_folder, f'resume_scorecard_{str(current_time)}.xlsx')

    scorecard.to_excel(file_path, index=False)
    # scorecard.to_excel(f'resume_scorecard{str(current_time)}.xlsx', index=False)
    print(f"Score Results saved to {file_path}")

# # Main Execution Flow
# if __name__ == "__main__":
#     folder_path = select_folder()
#     print(f"The resumes are selected from :- {folder_path}")
#     job_description = input("Please enter JOB description: ")

#     resume_df = extract_text_from_files(folder_path)
#     resume_df = process_resumes(resume_df, job_description)
#     save_results(resume_df)


import threading
from queue import Queue

num_threads = os.cpu_count()

# Threaded function to read files
def threaded_file_reader(file_queue, data):
    while not file_queue.empty():
        file_path = file_queue.get()
        filename = os.path.basename(file_path)

        try:
            print(f"Reading file: {filename}")  # Print the name of the file being processed
            if filename.endswith('.pdf'):
                text = read_pdf(file_path)
            elif filename.endswith('.docx'):
                text = read_docx(file_path)
            elif filename.endswith('.doc'):
                text = read_doc(file_path)
            elif filename.endswith('.txt'):
                text = read_txt(file_path)
            else:
                continue

            data.append({'resume_file_name': filename, 'resume_file_text': text})
        except Exception as e:
            print(f"Error processing file {filename}: {e}")
        finally:
            file_queue.task_done()

def extract_text_from_files(folder_path):
    file_queue = Queue()
    data = []

    # Populate the queue with file paths
    for filename in os.listdir(folder_path):
        file_queue.put(os.path.join(folder_path, filename))

    # Start threads
    threads = []
    for _ in range(num_threads):  # Adjust the number of threads based on system capacity
        thread = threading.Thread(target=threaded_file_reader, args=(file_queue, data))
        thread.start()
        threads.append(thread)

    # Wait for all threads to finish
    file_queue.join()
    for thread in threads:
        thread.join()

    return pd.DataFrame(data)

# Threaded function to process resumes
def threaded_resume_processor(resume_queue, job_description, results):
    conversation_jd = get_conversation_openai(TEMPLATES["job_description"])
    conversation_resume = get_conversation_openai(TEMPLATES["resume"])
    conversation_score = get_conversation_openai(TEMPLATES["score"])

    # Process job description once
    processed_jd = conversation_jd({"job_description_text": job_description})

    while not resume_queue.empty():
        try:
            index, resume = resume_queue.get()
            resume_text = resume["resume_file_text"]
            filename = resume["resume_file_name"]
            print(f"Processing resume: {filename}")  # Print the name of the resume being processed

            # Extract key aspects
            resume_response = conversation_resume({"resume_text": resume_text})

            # Score resume
            score_response = conversation_score({
                "resume_text": resume_text,
                "job_description": processed_jd
            })

            results[index] = {
                "resume_key_aspect": resume_response,
                "resume_score": score_response
            }
        except Exception as e:
            print(f"Error processing resume: {e}")
        finally:
            resume_queue.task_done()

def process_resumes(resume_df, job_description):
    resume_queue = Queue()
    results = {}

    # Populate the queue with resume data
    for i, resume in resume_df.iterrows():
        resume_queue.put((i, resume))

    # Start threads
    threads = []
    for _ in range(num_threads):  # Adjust the number of threads
        thread = threading.Thread(target=threaded_resume_processor, args=(resume_queue, job_description, results))
        thread.start()
        threads.append(thread)

    # Wait for all threads to finish
    resume_queue.join()
    for thread in threads:
        thread.join()

    # Update the DataFrame with the results
    for index, result in results.items():
        resume_df.loc[index, "resume_key_aspect"] = result["resume_key_aspect"]
        resume_df.loc[index, "resume_score"] = result["resume_score"]

    return resume_df

# Main Execution Flow
if __name__ == "__main__":
    folder_path = select_folder()
    print(f"The resumes are selected from: {folder_path}")
    job_description = input("Please enter JOB description: ")

    resume_df = extract_text_from_files(folder_path)
    resume_df = process_resumes(resume_df, job_description)
    save_results(resume_df)
