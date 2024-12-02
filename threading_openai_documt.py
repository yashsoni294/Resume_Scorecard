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
import threading
from queue import Queue

# Load API Key
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

openai.api_key = API_KEY


TEMPLATES = {
    "job_description" : """
        
        The text below is a job description:
        {job_description_text}

        Your task is to analyze the job description and extract critical aspects to evaluate a candidate's suitability effectively. Organize the extracted information into structured categories as outlined below. Ensure conciseness and avoid including assumptions or unnecessary details. The structured output will form the foundation for precise scoring.

        1. Candidate Profile
            1.1 Job-Related Keywords:

            Extract highly relevant keywords and phrases, focusing on essential skills, tools, technologies, and qualifications.
            Highlight terms that frequently appear, emphasizing the primary focus areas for the role.
            1.2 Relevant Past Roles and Responsibilities:

            Identify specific roles (e.g., "Project Manager," "Data Analyst") and responsibilities directly relevant to the role described.
            Highlight areas where past experience aligns with the position key objectives.
            1.3 Actionable Responsibilities:

            List clear, measurable, and action-oriented expectations (e.g., "Design and implement X system," "Lead Y project team") to define success in this role.
        2. Experience Requirements
            2.1 Years of Experience:

            Specify the required and preferred experience levels, distinguishing between mandatory and desirable years in relevant fields.
            2.2 Technical Skills:

            Provide a categorized list of required and preferred technical skills, specifying domain-specific tools, programming languages, platforms, or methodologies.
            Highlight core skills critical for the role versus those that are supplementary.
            2.3 Soft Skills and Interpersonal Abilities:

            List required soft skills (e.g., leadership, problem-solving) and interpersonal abilities (e.g., teamwork, collaboration).
            Include any role-specific examples mentioned, such as "strong stakeholder communication skills."

        3. Educational Qualifications and Certifications
            3.1 Minimum Educational Qualifications:

            State the explicit educational requirements for the role (e.g., "Bachelors degree in Computer Science").
            Differentiate between mandatory and preferred qualifications.
            3.2 Certifications and Specialized Training:

            Highlight certifications, licenses, or training programs required or preferred (e.g., "PMP certification," "AWS Certified Solutions Architect").
            Include both general and domain-specific certifications if applicable.

        Output Format:
        Organize the extracted information in bullet-point format under the categories listed above. Ensure the content is:

        Directly aligned with the job description.
        Actionable and structured to facilitate accurate scoring.
        Free from redundant details or assumptions.
        """ ,

    "resume" : """
        The text below is a resume:
        {resume_text}

        Your task is to extract critical information from the resume, focusing only on its content without making assumptions or adding external details. The extracted details should be structured, concise, and actionable to support effective scoring. Use the categories below for organization:

        1. Candidate Profile
            1.1 Keywords Identified:

            Extract relevant keywords that reflect the candidate's skills, roles, expertise, and domain knowledge.
            Highlight recurring themes or terms indicative of specialization or focus areas.
            1.2 Summary of Past Roles:

            Summarize the candidate primary roles, emphasizing key responsibilities and measurable achievements.
            Include details about progression or diversity in roles where mentioned (e.g., growth from Analyst to Manager).
            1.3 Measurable Achievements:

            Identify specific, quantifiable accomplishments (e.g., "Increased revenue by X%," "Reduced costs by Y%").
            Highlight the use of action-oriented language (e.g., "Led," "Implemented," "Designed").
        2. Experience Details

            2.1 Total Years of Experience:

            Indicate the total years of professional experience and the industries or domains the candidate has worked in.
            Include any explicit references to seniority (e.g., "5+ years in project management").
            2.2 Technical Skills and Proficiencies:

            Extract technical skills explicitly mentioned (e.g., tools, programming languages, platforms) and categorize them as core or supplementary.
            Include details of certifications or work examples that validate these skills.
            2.3 Soft Skills and Team Contributions:

            Highlight references to soft skills (e.g., problem-solving, adaptability) and team-related contributions (e.g., collaboration, mentoring).
            Focus on examples that demonstrate these abilities, such as leadership roles or cross-functional projects.
        3. Educational Qualifications and Certifications

            3.1 Educational Background:

            Note the highest qualification achieved, field of study, and any notable academic honors or achievements.
            Include additional qualifications that may complement the role.
            3.2 Certifications and Professional Training:

            List certifications, training programs, and licenses, specifying their relevance to the candidate's field or the role in question.
            Highlight certifications that indicate advanced expertise or specialization (e.g., "AWS Certified Solutions Architect").

        Output Format:
        Present the extracted details in the following format:

        Category: Subcategory/Point (e.g., Candidate Profile: Measurable Achievements).
        Use bullet points or short, clear sentences for each item.
        Ensure alignment with the resume content without adding interpretations or assumptions.

        """ , 
    "score" : """
        Your task is to evaluate the alignment between the provided resume and job description by analyzing three critical sections: Candidate Profile, Experience, and Educational Qualifications and Certifications. Based on your evaluation, assign a final score between 0 and 100, reflecting the overall suitability of the candidate for the job.

        Inputs:
        Resume Text:
        {resume_text}

        Job Description Text:
        {job_description}

        Scoring Guidelines:
        Evaluate the resume against the job description using the criteria outlined below. Assign marks in each category, calculate the total, and round the final score to the nearest whole number.

        1. Candidate Profile (Max 15 Marks)
            1.1 Job-Related Keywords (Max 5 Marks):
                5 Points: Resume includes all highly relevant keywords, indicating strong alignment with job requirements.
                3 Points: Resume includes many relevant keywords but misses some critical ones.
                1 Points: Resume includes few relevant keywords or misses key terms.
            1.2 Relevance of Past Roles to Job Description (Max 5 Marks):
                5 Points: Past roles and responsibilities strongly align with the job description.
                3 Points: Moderate alignment, with partial overlap in roles and responsibilities.
                1 Points: Limited relevance or weak alignment.
            1.3 Clarity of Responsibilities (Max 5 Marks):
                5 Points: Responsibilities are clearly defined using action words (e.g., "Developed," "Managed") with measurable outcomes.
                3 Points: Responsibilities are described but lack clear action words or measurable outcomes.
                1 Points: Responsibilities are vague or generic.

        2. Experience Section (Max 65 Marks)
            2.1 Years of Experience (Max 15 Marks):
                15 Points: Meets or exceeds the required years of experience.
                10 Points: Slightly below the required years but with relevant experience.
                5 Points: Limited relevance or inadequate years of experience.
            2.2 Matching Technical Skills (Max 40 Marks):
                40 Points: All technical skills mentioned in the job description are evident, supported by examples or certifications.
                25 Points: Most technical skills are evident, but examples or certifications are missing.
                15 Points: Some technical skills align, but several are missing.
                5 Points: Minimal or no alignment with the required technical skills.
            2.3 Communication and Teamwork (Max 10 Marks):
                10 Points: Strong evidence of soft skills, supported by examples (e.g., "Led a team of 5," "Facilitated cross-department collaboration").
                7 Points: Mentions soft skills but lacks specific examples.
                3 Points: Minimal or generic mention of soft skills.

        3. Educational Qualifications and Certifications (Max 20 Marks)
            3.1 Minimum Educational Qualifications (Max 15 Marks):
                15 Points: Meets or exceeds the educational qualifications specified in the job description.
                10 Points: Meets basic qualifications but lacks advanced or preferred qualifications.
                5 Points: Does not fully meet the educational qualifications.
            3.2 Additional Certifications/Training Programs (Max 5 Marks):
                5 Points: Certifications/training are directly relevant to the job description (e.g., industry-specific certifications).
                3 Points: Certifications or training are partially relevant to the job description.
                1 Point: No additional certifications or irrelevant certifications.

        Additional Refinements:
            Ensure that scoring accounts for both the breadth and depth of alignment between the resume and job description.
            Emphasize evidence-backed qualifications and experience to avoid scoring inflated or unsupported claims.
        Output:
            Provide the final calculated score as a single whole number (0 – 100) with no additional explanation or text. If you are not able to score the resume then you can give 0 score to the resume
        """

}


# Initialize DataFrame
resume_df = pd.DataFrame(columns=['resume_file_name', 'resume_file_text', 'resume_key_aspect', 'resume_score'])


def get_conversation_openai(template, model="gpt-4o-mini", temperature=0.1, max_tokens=None):

    """
    Creates a function that interacts with the OpenAI model based on a provided template.
    
    Args:
        template (str): A string template for generating prompts dynamically. The template should use placeholders
                        that can be filled with dynamic input values.
        model (str, optional): The name of the OpenAI model to use for generating responses. Defaults to "gpt-4o-mini".
        temperature (float, optional): Sampling temperature to control the randomness of the response. 
                                       Lower values make output more focused and deterministic. Defaults to 0.1.
        max_tokens (int, optional): The maximum number of tokens to include in the response. Defaults to None, 
                                    allowing the model to determine the length.

    Returns:
        function: A callable function that takes a dictionary of inputs, formats the prompt based on the template,
                  and interacts with the OpenAI model to generate a response.

    """
 
    # Define a nested function to handle API interaction
    def call_openai_model(inputs):
        """
        Invokes the OpenAI model using the provided template and inputs.
        
        Args:
            inputs (dict): A dictionary containing values for the placeholders in the template.

        Returns:
            str: The content of the response generated by the OpenAI model.
        """
        # Generate the prompt by formatting the template with the provided inputs
        prompt = PromptTemplate.from_template(template).format(**inputs)
        # Call the OpenAI Chat API to generate a response
        response = openai.ChatCompletion.create(
            model=model,
            messages=[{"role": "system", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens
        )
        # Extract and return the content of the response
        return response["choices"][0]["message"]["content"]
    
    # Return the nested function for reuse
    return call_openai_model

def select_folder():
    """
    Opens a folder selection dialog and returns the selected folder path.
    
    This function uses the `tkinter` library to create a folder selection dialog. 
    The dialog allows users to browse and select a folder. The returned path uses 
    Windows-style backslashes for compatibility with Windows file systems.

    Returns:
        str: The path of the selected folder, formatted with backslashes. 
             If the user cancels the dialog, an empty string is returned.
    """
    # Create a hidden root window for the dialog
    root = tk.Tk()
    root.withdraw() # Hide the root window

    # Open the folder selection dialog and get the selected folder path
    folder_path = filedialog.askdirectory(title="Select Folder")

    # Replace forward slashes with backslashes for Windows compatibility
    return folder_path.replace('/', '\\')

def read_pdf(file_path):
    """
    Reads and extracts text from a PDF file.
    
    This function uses the `PyPDF2` library to read a PDF file and extract text 
    from each page. In case of an error (e.g., unsupported PDF format or file corruption), 
    it prints an error message and returns an empty string.

    Args:
        file_path (str): The full path to the PDF file to be read.

    Returns:
        str: The extracted text from the PDF file. If an error occurs, an empty string is returned.
    """
    # Initialize an empty string to store the extracted text
    text = ""
    try:
        # Open the PDF file in binary read mode
        with open(file_path, 'rb') as file:
            # Initialize the PDF reader
            reader = PyPDF2.PdfReader(file)
            # Iterate through all pages and extract text
            for page in reader.pages:
                text += page.extract_text()
    # Handle exceptions and print an error message if reading fails
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
    # Return the extracted text (or an empty string if an error occurred)
    return text

def read_txt(file_path):
    """
    Reads and extracts text from a TXT file.
    
    This function reads the contents of a text file using UTF-8 encoding. If an error occurs 
    (e.g., file not found, encoding issues), it prints an error message and returns an empty string.

    Args:
        file_path (str): The full path to the TXT file to be read.

    Returns:
        str: The text content of the TXT file. If an error occurs, an empty string is returned.

    """
    # Initialize an empty string to store the text
    text = ""
    try:
        # Open the file in read mode with UTF-8 encoding
        with open(file_path, 'r', encoding='utf-8') as file:
            # Read the entire file content
            text = file.read()
    # Handle exceptions and print an error message if reading fails
    except Exception as e:
        print(f"Error reading TXT file {file_path}: {e}")
    # Return the file content (or an empty string if an error occurred)
    return text

def read_docx(file_path):
    """
    Reads and extracts text from a DOCX file.
    
    This function uses the `python-docx` library to read the content of a DOCX file. 
    It extracts text from all paragraphs in the document. If an error occurs (e.g., 
    file corruption or unsupported format), it prints an error message and returns an empty string.

    Args:
        file_path (str): The full path to the DOCX file to be read.

    Returns:
        str: The text content of the DOCX file, with paragraphs separated by newlines. 
             If an error occurs, an empty string is returned.
    """

    try:
        # Open the DOCX file using python-docx
        doc = docx.Document(file_path)
        # Extract text from all paragraphs and join them with newlines
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)
    except Exception as e:
        # Handle exceptions and print an error message if reading fails and returns a empty string.
        print(f"Error reading DOCX {file_path}: {e}")
        return ""

def read_doc(file_path):
    """
    Reads and extracts text from a DOC (Microsoft Word 97-2003) file.
    
    This function uses the `pywin32` library to interact with the Microsoft Word application 
    through COM automation. It extracts the text content of the document. If an error occurs 
    (e.g., Word not installed, file corruption), it prints an error message and returns an empty string.

    Args:
        file_path (str): The full path to the DOC file to be read.

    Returns:
        str: The text content of the DOC file. If an error occurs, an empty string is returned.
    """

    try:
        # Use pywin32 to create a Word application instance
        word = win32.Dispatch("Word.Application")
        word.Visible = False # Keep the Word application hidden

        # Open the DOC file
        doc = word.Documents.Open(file_path)

        # Extract the text content of the document
        text = doc.Content.Text

        # Close the document and quit the Word application
        doc.Close(False)
        word.Quit()

        # Return the extracted text
        return text
    # Handle exceptions and print an error message if reading fails and returns empty string.
    except Exception as e:
        print(f"Error reading DOC {file_path}: {e}")
        return ""


def extract_text_from_files(folder_path):
    """
    Extracts text from various file types in a specified folder and organizes it into a DataFrame.
    
    This function processes files in a folder with supported extensions (.pdf, .docx, .doc, .txt). 
    It uses the respective functions (`read_pdf`, `read_docx`, `read_doc`, `read_txt`) to extract 
    text and compiles the results into a pandas DataFrame. Unsupported file types are skipped.

    Args:
        folder_path (str): The path to the folder containing the files to process.

    Returns:
        pd.DataFrame: A DataFrame with columns:
            - 'resume_file_name': The name of the file.
            - 'resume_file_text': The extracted text content of the file.
    """
    # Initialize a list to store data for each file
    data = []

    # Iterate through unique files in the specified folder
    for filename in set(os.listdir(folder_path)):
        # Build the full file path
        file_path = os.path.join(folder_path, filename)
        # Extract text based on the file extension
        if filename.endswith('.pdf'):
            text = read_pdf(file_path)
        elif filename.endswith('.docx'):
            text = read_docx(file_path)
        elif filename.endswith('.doc'):
            text = read_doc(file_path)
        elif filename.endswith('.txt'):
            text = read_txt(file_path)
        else:
            # Skip unsupported file types
            continue
        # Append the filename and extracted text to the data list
        data.append({'resume_file_name': filename, 'resume_file_text': text})
    # Convert the data into a pandas DataFrame and returns.
    return pd.DataFrame(data)

def process_resumes(resume_df, job_description):
    """
    Processes a DataFrame of resumes by extracting key aspects and scoring them 
    based on a given job description.

    This function uses OpenAI models to:
    1. Process the job description to identify its key aspects.
    2. Extract key aspects from each resume.
    3. Score each resume based on its relevance to the processed job description.

    Args:
        resume_df (pd.DataFrame): A DataFrame containing resumes with columns:
            - 'resume_file_text': The text content of each resume.
            - 'resume_file_name': The name of each resume file.
        job_description (str): The text of the job description to analyze.

    Returns:
        pd.DataFrame: The updated DataFrame with additional columns:
            - 'resume_key_aspect': Key aspects extracted from each resume.
            - 'resume_score': Relevance score of each resume to the job description.
    """

    # Initialize the OpenAI conversation for job description processing
    conversation_jd = get_conversation_openai(TEMPLATES["job_description"])
    jd_response = conversation_jd({"job_description_text": job_description})
    processed_jd = jd_response # Store the processed job description

    # Initialize OpenAI conversations for resume processing and scoring
    conversation_resume = get_conversation_openai(TEMPLATES["resume"])
    conversation_score = get_conversation_openai(TEMPLATES["score"])

    # Iterate through each resume in the DataFrame
    for i in range(len(resume_df)):
        # Extract the text content of the current resume
        resume_text = resume_df.loc[i, "resume_file_text"]
        
        # Extract key aspects from the resume
        resume_response = conversation_resume({"resume_text": resume_text})
        resume_df.loc[i, "resume_key_aspect"] = resume_response
        
        # Score the resume based on the processed job description
        score_response = conversation_score({
            "resume_text": resume_text,
            "job_description": processed_jd
        })
        resume_df.loc[i, "resume_score"] = score_response

        # Log progress to the console
        print(f"{i+1}. Working on - ",resume_df["resume_file_name"][i])

    # Return the updated DataFrame
    return resume_df

def save_results(resume_df):
    """
    Saves the resume scores to an Excel file in the user's Downloads folder.
    
    This function sorts the resumes by their scores in descending order, extracts the 
    relevant columns, and saves the results to an Excel file in the Downloads folder. 
    The file is timestamped to ensure uniqueness.

    Args:
        resume_df (pd.DataFrame): A DataFrame containing resume data with at least the columns:
            - 'resume_file_name': The name of each resume file.
            - 'resume_score': The score assigned to each resume.

    Returns:
        None
    """
    # Sort resumes by their scores in descending order
    resume_df.sort_values(by='resume_score', ascending=False, inplace=True)

    # Extract only the relevant columns for the scorecard
    scorecard = resume_df[["resume_file_name", "resume_score"]]

    # Determine the path to the user's Downloads folder
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

    # Get the current date and time in a formatted string
    current_time = datetime.now().strftime("%d-%b-%Y_%I-%M-%S_%p")

    # Define the full file path for the output Excel file
    file_path = os.path.join(downloads_folder, f'resume_scorecard_{str(current_time)}.xlsx')

    # Save the scorecard to an Excel file
    scorecard.to_excel(file_path, index=False)
    
    # Print a confirmation message with the file path
    print(f"Score Results saved to {file_path}")


num_threads = os.cpu_count()

def threaded_resume_processor(resume_queue, job_description, results):
    """
    Processes resumes from a queue in a threaded environment, extracting key aspects 
    and calculating scores based on a given job description.

    This function uses OpenAI models to:
    1. Process the job description to extract its key aspects.
    2. Extract key aspects from each resume.
    3. Calculate a relevance score for each resume against the job description.

    Args:
        resume_queue (queue.Queue): A thread-safe queue containing resumes to process. 
            Each item in the queue is a tuple (index, resume_dict), where:
                - index (int): The position of the resume in the original DataFrame.
                - resume_dict (dict): A dictionary with keys:
                    - 'resume_file_name': The name of the resume file.
                    - 'resume_file_text': The text content of the resume.
        job_description (str): The text of the job description.
        results (dict): A thread-safe dictionary to store the processed results. 
            Each entry in the dictionary is keyed by the index and contains:
                - 'resume_key_aspect': Key aspects extracted from the resume.
                - 'resume_score': Relevance score for the resume.

    Returns:
        None
    """
    # Initialize OpenAI conversations
    conversation_jd = get_conversation_openai(TEMPLATES["job_description"])
    conversation_resume = get_conversation_openai(TEMPLATES["resume"])
    conversation_score = get_conversation_openai(TEMPLATES["score"])

    # Process job description once
    processed_jd = conversation_jd({"job_description_text": job_description})

    # Process resumes from the queue
    while not resume_queue.empty():
        try:
            index, resume = resume_queue.get()
            resume_text = resume["resume_file_text"]
            filename = resume["resume_file_name"]
            print(f"Processing resume: {filename}")  # Log the current resume being processed

            # Extract key aspects from the resume
            resume_response = conversation_resume({"resume_text": resume_text})

            # Calculate the resume score based on the processed job description
            score_response = conversation_score({
                "resume_text": resume_text,
                "job_description": processed_jd
            })

            # Store the results in the thread-safe dictionary
            results[index] = {
                "resume_key_aspect": resume_response,
                "resume_score": score_response
            }
        except Exception as e:
            # Log any errors encountered during processing
            print(f"Error processing resume: {e}")
        finally:
            # Mark the task as done, signaling to the queue
            resume_queue.task_done()

def process_resumes(resume_df, job_description):
    """
    Processes resumes in parallel using multiple threads, extracting key aspects 
    and calculating scores based on a given job description.

    This function populates a queue with resumes, then processes each resume 
    in a threaded manner using the `threaded_resume_processor`. After all threads 
    have finished processing, it updates the DataFrame with the extracted key aspects 
    and scores for each resume.

    Args:
        resume_df (pd.DataFrame): A DataFrame containing resumes with at least the columns:
            - 'resume_file_text': The text content of each resume.
            - 'resume_file_name': The name of each resume file.
        job_description (str): The text of the job description to analyze.

    Returns:
        pd.DataFrame: The updated DataFrame with additional columns:
            - 'resume_key_aspect': Key aspects extracted from each resume.
            - 'resume_score': Relevance score of each resume to the job description.
    """

    # Initialize the queue and results dictionary
    resume_queue = Queue()
    results = {}

    # Populate the queue with resume data from the DataFrame
    for i, resume in resume_df.iterrows():
        resume_queue.put((i, resume))

    # Number of threads to use for processing
    threads = []
    for _ in range(num_threads):  # Adjust the number of threads
        thread = threading.Thread(target=threaded_resume_processor, args=(resume_queue, job_description, results))
        thread.start()
        threads.append(thread)

    # Wait for all threads to finish processing
    resume_queue.join()
    for thread in threads:
        thread.join()

    # Update the DataFrame with the results from the threads
    for index, result in results.items():
        resume_df.loc[index, "resume_key_aspect"] = result["resume_key_aspect"]
        resume_df.loc[index, "resume_score"] = result["resume_score"]

    # Return the updated DataFrame with scores and key aspects
    return resume_df

# Main Execution Flow
if __name__ == "__main__":
    folder_path = select_folder()
    print(f"The resumes are selected from: {folder_path}")
    job_description = input("Please enter JOB description: ")

    resume_df = extract_text_from_files(folder_path)
    resume_df = process_resumes(resume_df, job_description)
    save_results(resume_df)
