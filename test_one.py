import os
from dotenv import load_dotenv
import openai
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
from langchain.prompts import PromptTemplate
from datetime import datetime
import threading
from queue import Queue

# Load environment variables
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")
openai.api_key = API_KEY

# Constants
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
            Provide the final calculated score as a single whole number (0 – 100) with no additional explanation or text.
            Never provide text explanation as Output give 0 Score Instead.
        """

}

NUM_THREADS = os.cpu_count()


def initialize_openai_prompt(template, model="gpt-4o-mini", temperature=0.1, max_tokens=None):
    """Initializes an OpenAI completion function based on a template."""
    def call_model(inputs):
        prompt = PromptTemplate.from_template(template).format(**inputs)
        response = openai.ChatCompletion.create(
            model=model,
            messages=[{"role": "system", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return response["choices"][0]["message"]["content"]
    return call_model


def select_folder():
    """Prompts the user to select a folder."""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder").replace('/', '\\')


def read_pdf(file_path):
    """Reads text from a PDF file."""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            return ''.join(page.extract_text() for page in reader.pages)
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return ""


def read_txt(file_path):
    """Reads text from a TXT file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        print(f"Error reading TXT {file_path}: {e}")
        return ""


def read_docx(file_path):
    """Reads text from a DOCX file."""
    try:
        doc = docx.Document(file_path)
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return ""


def read_doc(file_path):
    """Reads text from a DOC file."""
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
    """Extracts text from all files in a folder using multithreading."""
    file_queue = Queue()
    data = []

    for filename in os.listdir(folder_path):
        file_queue.put(os.path.join(folder_path, filename))

    def threaded_reader(queue, data_list):
        while not queue.empty():
            file_path = queue.get()
            filename = os.path.basename(file_path)
            try:
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
                data_list.append({'resume_file_name': filename, 'resume_file_text': text})
            except Exception as e:
                print(f"Error processing file {filename}: {e}")
            finally:
                queue.task_done()

    threads = [threading.Thread(target=threaded_reader, args=(file_queue, data)) for _ in range(NUM_THREADS)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    return pd.DataFrame(data)


def process_resumes(resume_df, job_description):
    """Processes resumes to extract key aspects and scores."""
    jd_processor = initialize_openai_prompt(TEMPLATES["job_description"])
    resume_processor = initialize_openai_prompt(TEMPLATES["resume"])
    score_processor = initialize_openai_prompt(TEMPLATES["score"])

    processed_jd = jd_processor({"job_description_text": job_description})

    def threaded_processor(queue, results):
        while not queue.empty():
            index, resume = queue.get()
            try:
                resume_text = resume["resume_file_text"]
                key_aspect = resume_processor({"resume_text": resume_text})
                score = score_processor({"resume_text": resume_text, "job_description": processed_jd})
                results[index] = {"resume_key_aspect": key_aspect, "resume_score": score}
            except Exception as e:
                print(f"Error processing resume {resume['resume_file_name']}: {e}")
            finally:
                queue.task_done()

    resume_queue = Queue()
    for i, row in resume_df.iterrows():
        resume_queue.put((i, row))

    results = {}
    threads = [threading.Thread(target=threaded_processor, args=(resume_queue, results)) for _ in range(NUM_THREADS)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    for index, result in results.items():
        resume_df.loc[index, "resume_key_aspect"] = result["resume_key_aspect"]
        resume_df.loc[index, "resume_score"] = result["resume_score"]

    return resume_df


def save_results(resume_df):
    """Saves results to an Excel file."""
    resume_df.sort_values(by='resume_score', ascending=False, inplace=True)
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    timestamp = datetime.now().strftime("%d-%b-%Y_%I-%M-%S_%p")
    file_path = os.path.join(downloads_folder, f'resume_scorecard_{timestamp}.xlsx')
    resume_df.to_excel(file_path, index=False)
    print(f"Results saved to {file_path}")


if __name__ == "__main__":
    folder_path = select_folder()
    print(f"The resumes are selected from: {folder_path}")
    job_description = input("Please enter the job description: ")

    resumes = extract_text_from_files(folder_path)
    processed_resumes = process_resumes(resumes, job_description)
    save_results(processed_resumes)
