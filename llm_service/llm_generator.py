
import os
import requests
import json
from openai import OpenAI
from dotenv import load_dotenv
import base64

load_dotenv()

# Retrieve API keys
HUGGINGFACE_API_KEY = os.getenv("HUGGINGFACE_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")  # Hypothetical


# Function to encode the image
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


# If using the official OpenAI Python library:




# If using Anthropic's Python library for Claude (hypothetical usage):
#   pip install anthropic
#   import anthropic
#   anthropic.Client(ANTHROPIC_API_KEY)

def generate_llm_response(prompt, provider="openai", model="gpt-4o", temperature=0.7):
    """
    Generates a response from various LLM providers (OpenAI, Hugging Face, Claude, Google Gemini).
    
    :param prompt: The prompt or query string.
    :param provider: Which LLM provider to use ('openai', 'huggingface', 'claude', 'gemini').
    :param model: Model name (e.g., 'gpt-4', 'gpt-4o', 'claude-v1', 'google-gemini', etc.).
    :param temperature: Sampling temperature (if applicable).
    :return: The text response from the LLM, or an error string if something fails.
    """
    try:
        if provider.lower() == "openai":
            # Using OpenAI's official Python library
            client = OpenAI(api_key=OPENAI_API_KEY)
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=temperature,
            )
            return response.choices[0].message.content
        
        elif provider.lower() == "huggingface":
            # Using Hugging Face Inference API
            # Make sure to have HUGGINGFACE_API_KEY set in your environment
            # and set your model endpoint, e.g., "bigscience/bloomz"
            huggingface_url = f"https://api-inference.huggingface.co/models/{model}"
            headers = {"Authorization": f"Bearer {HUGGINGFACE_API_KEY}"}
            
            payload = {
                "inputs": prompt,
                "parameters": {"temperature": temperature, "max_new_tokens": 300},
                "options": {"wait_for_model": True}
            }
            
            hf_response = requests.post(huggingface_url, headers=headers, json=payload)
            if hf_response.status_code == 200:
                data = hf_response.json()
                # Some Hugging Face models return a list of generated texts
                # You may need to adapt parsing logic for your specific model
                if isinstance(data, list) and len(data) > 0 and "generated_text" in data[0]:
                    return data[0]["generated_text"]
                else:
                    return str(data)
            else:
                return f"HuggingFace API Error: {hf_response.text}"
        
        elif provider.lower() == "claude":
            # Using Anthropic's API for Claude
            # The code snippet below is illustrative; official usage may differ
            # https://github.com/anthropics/anthropic-sdk-python
            # import anthropic
            # client = anthropic.Client(api_key=ANTHROPIC_API_KEY)
            
            # Hypothetical usage:
            # conversation = anthropic.ChatCompletion(
            #     model=model,
            #     messages=[{"role": "user", "content": prompt}],
            #     temperature=temperature
            # )
            # return conversation["completion"]
            
            # We'll simulate a direct request for demonstration:
            # (Replace with actual library usage or requests to Claude's endpoint.)
            
            claude_url = "https://api.anthropic.com/v1/complete"  # Example endpoint
            headers = {"x-api-key": ANTHROPIC_API_KEY, "Content-Type": "application/json"}
            data = {
                "model": model,
                "prompt": f"\n\nHuman: {prompt}\n\nAssistant:",
                "max_tokens_to_sample": 300,
                "temperature": temperature,
            }
            
            claude_response = requests.post(claude_url, headers=headers, json=data)
            if claude_response.status_code == 200:
                res_json = claude_response.json()
                # The exact response structure depends on Anthropic's API
                if "completion" in res_json:
                    return res_json["completion"]
                else:
                    return str(res_json)
            else:
                return f"Claude API Error: {claude_response.text}"
        
        elif provider.lower() == "gemini":
            # Hypothetical usage for Google Gemini
            # There's currently no official Python library or public endpoint for Gemini at time of writing.
            # Below is a placeholder to illustrate usage.
            
            # Example usage with PaLM or hypothetical Gemini endpoint:
            gemini_url = f"https://generativelanguage.googleapis.com/v1beta2/models/{model}:generateText"
            headers = {
                "Authorization": f"Bearer {GEMINI_API_KEY}",
                "Content-Type": "application/json"
            }
            data = {
                "prompt": {"text": prompt},
                "temperature": temperature,
                "candidate_count": 1
            }
            gemini_response = requests.post(gemini_url, headers=headers, json=data)
            if gemini_response.status_code == 200:
                res_json = gemini_response.json()
                # Hypothetical response structure
                if "candidates" in res_json and len(res_json["candidates"]) > 0:
                    return res_json["candidates"][0].get("output", "")
                else:
                    return str(res_json)
            else:
                return f"Gemini API Error: {gemini_response.text}"
        
        else:
            return "LLM Error: Unknown provider specified."
    
    except Exception as e:
        return f"LLM Error: {str(e)}"
    
    
def generate_image_description(image_path, prompt,provider="openai", model="gpt-4o-mini",temperature=0.7):
    """
    Generates an image description using OpenAI's API.
    Since OpenAI's API does not accept image binary directly for captioning,
    the image is assumed to be available at a public URL.
    
    :param image_path: Path to the input image.
    :param provider: LLM provider, default 'openai'.
    :param model: LLM model name.
    :param temperature: Sampling temperature.
    :return: A generated caption describing the image.
    """
    client = OpenAI(api_key=OPENAI_API_KEY) 
    
    image_path = image_path
    base64_image = encode_image(image_path)
    response = client.chat.completions.create(
                                            model=model,
                                            messages=[
                                                                    {
                                                                        "role": "user",
                                                                        "content": [
                                                                            {
                                                                                "type": "text",
                                                                                "text": prompt,
                                                                            },
                                                                            {
                                                                                "type": "image_url",
                                                                                "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"},
                                                                            },
                                                                        ],
                                                                    }
                                                                ],
                                                            )



    return response.choices[0].message.content



def generate_llm_json(prompt,event,provider="openai", model="gpt-4o-2024-08-06",temperature=0.7):
    try:
        if provider.lower() == "openai":
            client = OpenAI(api_key=OPENAI_API_KEY)
            completion = client.beta.chat.completions.parse(
            model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=temperature,
            response_format=event,
            )
            return completion.choices[0].message.parsed
    except Exception as e:
        return f"LLM Error: {str(e)}"
    
