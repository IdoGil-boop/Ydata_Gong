import anthropic
import os
from typing import List, Dict, Any

def setup_anthropic_client():
    """
    Sets up and returns Anthropic client using API key from environment variables
    """
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("Anthropic API key not found in environment variables")
    client = anthropic.Anthropic(api_key=api_key)
    return client

def get_claude_completion(
    messages: List[Dict[str, str]], 
    model: str = "claude-3-opus-20240229",
    temperature: float = 0.7,
    max_tokens: int = 1000
) -> str:
    """
    Gets a completion response from Anthropic's Claude API

    Args:
        messages: List of message dictionaries with 'role' and 'content' keys
        model: Anthropic model to use
        temperature: Controls randomness (0-1)
        max_tokens: Maximum tokens in response

    Returns:
        Response text from the model
    """
    client = setup_anthropic_client()
    
    try:
        # Convert OpenAI-style messages to Anthropic format
        anthropic_messages = []
        for message in messages:
            role = "assistant" if message["role"] == "assistant" else "user"
            anthropic_messages.append({"role": role, "content": message["content"]})
        
        response = client.messages.create(
            model=model,
            messages=anthropic_messages,
            temperature=temperature,
            max_tokens=max_tokens
        )
        return response.content[0].text
    except Exception as e:
        print(f"Error getting Claude completion: {str(e)}")
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
    
    return get_claude_completion(messages)

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
    
    return get_claude_completion(messages)

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
    
    return get_claude_completion(messages)
