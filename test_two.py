import openai
from langchain.prompts import PromptTemplate
import os
from dotenv import load_dotenv
# Load API Key
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

openai.api_key = API_KEY

def get_conversation_openai(template, model="gpt-4o-mini", temperature=0.1, max_tokens=None):
    """
    Initializes a conversation using a template and OpenAI's GPT-4 model integration.
    
    Args:
        template (str): Template content for the conversation.
        model (str): OpenAI model to use. Default is 'gpt-4'.
        temperature (float): Sampling temperature. Default is 0.1.
        max_tokens (int, optional): Maximum number of tokens in the output. Default is None.
    
    Returns:
        Callable: Configured conversation object.
    """
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

# Example usage
if __name__ == "__main__":
    # Define your template
    TEMPLATES = {
        "job_description": "Analyze the following job description: {job_description_text}"
    }
    
    job_description = "Looking for a Python developer with experience in AI and machine learning."
    
    conversation_jd = get_conversation_openai(TEMPLATES["job_description"])
    jd_response = conversation_jd({"job_description_text": job_description})
    print(jd_response)