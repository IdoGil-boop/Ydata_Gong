import dspy
import json
from typing import List, Dict, Any
import os
import os.path


# Load environment variables from .env file if it exists
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("dotenv package not installed. Skipping .env loading.")

# Configure DSPy to use Nebius model
nebius_api_key = os.getenv("NEBIUS_API_KEY")
if not nebius_api_key:
    raise ValueError("NEBIUS_API_KEY environment variable not set")

nebius_model = dspy.Nebius(api_key=nebius_api_key, model="your-nebius-model-name")
dspy.settings.configure(lm=nebius_model)

# Set up prompt logging
os.makedirs("logs", exist_ok=True)
dspy.settings.configure(trace=["logs/dspy_prompts.log"])

class ExtractTextFromSlides(dspy.Signature):
    """Extract text from slides in JSON format."""
    images = dspy.InputField(desc="Images of slides")
    json_output = dspy.OutputField(desc="Extracted text in JSON format")

class CreatePlaceholders(dspy.Signature):
    """Replace all text with [PLACEHOLDER] in JSON format."""
    json_input = dspy.InputField(desc="JSON with extracted text")
    json_with_placeholders = dspy.OutputField(desc="JSON with placeholders")

class FillPlaceholders(dspy.Signature):
    """Fill placeholders with fitting text based on context."""
    context_rtf = dspy.InputField(desc="Context in RTF format")
    json_with_placeholders = dspy.InputField(desc="JSON with placeholders")
    filled_json = dspy.OutputField(desc="JSON with filled placeholders")

class CompareTexts(dspy.Signature):
    """Compare extracted text with filled text and create comparison table."""
    images = dspy.InputField(desc="Images of slides")
    original_json = dspy.InputField(desc="JSON with original content")
    filled_json = dspy.InputField(desc="JSON with filled content")
    comparison_table = dspy.OutputField(desc="Table comparing filled slides to actual slides with similarity scores")

# Implement the modules using DSPy
class ExtractTextModule(dspy.Module):
    def __init__(self):
        super().__init__()
        self.extract_text = dspy.ChainOfThought(ExtractTextFromSlides)
    
    def forward(self, images):
        result = self.extract_text(images=images)
        return result.json_output

class CreatePlaceholdersModule(dspy.Module):
    def __init__(self):
        super().__init__()
        self.create_placeholders = dspy.ChainOfThought(CreatePlaceholders)
    
    def forward(self, json_input):
        result = self.create_placeholders(json_input=json_input)
        return result.json_with_placeholders

class FillPlaceholdersModule(dspy.Module):
    def __init__(self):
        super().__init__()
        self.fill_placeholders = dspy.ChainOfThought(FillPlaceholders)
    
    def forward(self, context_rtf, json_with_placeholders):
        result = self.fill_placeholders(context_rtf=context_rtf, json_with_placeholders=json_with_placeholders)
        return result.filled_json

class CompareTextsModule(dspy.Module):
    def __init__(self):
        super().__init__()
        self.compare_texts = dspy.ChainOfThought(CompareTexts)
    
    def forward(self, images, original_json, filled_json):
        result = self.compare_texts(images=images, original_json=original_json, filled_json=filled_json)
        return result.comparison_table

# Create a pipeline to run all modules
class SlidePipeline:
    def __init__(self):
        self.extract_module = ExtractTextModule()
        self.placeholder_module = CreatePlaceholdersModule()
        self.fill_module = FillPlaceholdersModule()
        self.compare_module = CompareTextsModule()
    
    def extract_text(self, images):
        """Run Prompt 1: Extract text from slides"""
        return self.extract_module(images)
    
    def create_placeholders(self, json_input):
        """Run Prompt 2: Replace text with placeholders"""
        return self.placeholder_module(json_input)
    
    def fill_placeholders(self, context_rtf, json_with_placeholders):
        """Run Prompt 3: Fill placeholders with context"""
        return self.fill_module(context_rtf, json_with_placeholders)
    
    def compare_texts(self, images, original_json, filled_json):
        """Run Prompt 4: Compare extracted text with filled text"""
        return self.compare_module(images, original_json, filled_json)
    
    def run_full_pipeline(self, images, context_rtf):
        """Run the full pipeline from extraction to comparison"""
        original_json = self.extract_text(images)
        json_with_placeholders = self.create_placeholders(original_json)
        filled_json = self.fill_placeholders(context_rtf, json_with_placeholders)
        comparison = self.compare_texts(images, original_json, filled_json)
        
        return {
            "original_json": original_json,
            "json_with_placeholders": json_with_placeholders,
            "filled_json": filled_json,
            "comparison": comparison
        }
