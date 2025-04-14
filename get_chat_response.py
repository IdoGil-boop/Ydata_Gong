import openai
import os
from typing import List, Dict, Any

def setup_openai_client():
    """
    Sets up and returns OpenAI client using API key from environment variables
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OpenAI API key not found in environment variables")
    openai.api_key = api_key
    return openai

def get_chat_completion(
    messages: List[Dict[str, str]], 
    model: str = "gpt-3.5-turbo",
    temperature: float = 0.7,
    max_tokens: int = 1000
) -> str:
    """
    Gets a chat completion response from OpenAI's API

    Args:
        messages: List of message dictionaries with 'role' and 'content' keys
        model: OpenAI model to use
        temperature: Controls randomness (0-1)
        max_tokens: Maximum tokens in response

    Returns:
        Response text from the model
    """
    client = setup_openai_client()
    
    try:
        response = client.ChatCompletion.create(
            model=model,
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error getting chat completion: {str(e)}")
        return ""

def generate_slide_content(prompt: str = None, questions: List[str] = None) -> str:
    """
    Generates content for a slide based on either a prompt or oriented questions

    Args:
        prompt: Prompt describing the desired slide content
        questions: List of oriented questions to base content on

    Returns:
        Generated slide content
    """
    if prompt is not None:
        content = prompt
    elif questions is not None:
        content = f"""Generate slide content based on the following oriented questions:
            {questions}"""
    else:
        raise ValueError("Either prompt or questions must be provided")

    messages = [
        {"role": "system", "content": "You are a helpful presentation content creator."},
        {"role": "user", "content": content}
    ]
    
    return get_chat_completion(messages)

def improve_slide_content(current_content: str) -> str:
    """
    Improves existing slide content

    Args:
        current_content: Current content of the slide

    Returns:
        Improved slide content
    """
    prompt = f"""Please improve the following slide content to make it more engaging 
                and impactful while maintaining key information:
                
                {current_content}"""
                
    messages = [
        {"role": "system", "content": "You are an expert presentation editor."},
        {"role": "user", "content": prompt}
    ]
    
    return get_chat_completion(messages)

def generate_oriented_questions(slide_content: str) -> List[str]:
    """
    Generates oriented questions for a given slide content

    Args:
        slide_content: Content of the slide to generate questions for

    Returns:
        List of oriented questions
    """
    prompt = f"""Generate oriented questions for the following slide content:
        {slide_content}"""
    
    messages = [
        {"role": "system", "content": "You are a helpful presentation content creator."},
        {"role": "user", "content": prompt}
    ]
    
    return get_chat_completion(messages)