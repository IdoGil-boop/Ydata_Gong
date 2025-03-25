import pandas as pd
import pptx


# Function to replace text with PLACEHOLDER if it's wrapped in [[]], [] or {}
def replace_placeholders(presentation):
    for slide in presentation.slides:

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for paragraph in shape.text_frame.paragraphs:
                paragraph_text = paragraph.text

                # Check if the entire paragraph is a placeholder
                if (paragraph_text.strip().startswith('[') and paragraph_text.strip().endswith(']')) or \
                        (paragraph_text.strip().startswith('[[') and paragraph_text.strip().endswith(']]')) or \
                        (paragraph_text.strip().startswith('{') and paragraph_text.strip().endswith('}')):
                    # Clear existing runs
                    for idx in range(len(paragraph.runs)):
                        if idx == 0:
                            paragraph.runs[0].text = "PLACEHOLDER"
                        else:
                            paragraph.runs[idx].text = ""
                else:
                    # Check for placeholders within the text
                    import re
                    # Patterns for all three types of placeholders
                    placeholder_patterns = [
                        r'\[.*?\]',  # [text]
                        r'\[\[.*?\]\]',  # [[text]]
                        r'\{.*?\}'  # {text}
                    ]

                    # Process each run in the paragraph
                    for run in paragraph.runs:
                        run_text = run.text
                        if any(char in run_text for char in ['[', '{', '}']):
                            # Replace all instances of placeholders with PLACEHOLDER
                            modified_text = run_text
                            for pattern in placeholder_patterns:
                                modified_text = re.sub(pattern, 'PLACEHOLDER', modified_text)
                            run.text = modified_text


# For visualization, extract the text to show the changes
def extract_formatted_text(presentation):
    all_text_boxes = []

    for slide_index, slide in enumerate(presentation.slides):
        slide_text_boxes = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            # Process each paragraph in the text frame
            paragraphs_text = []
            current_text_block = []

            for paragraph in shape.text_frame.paragraphs:
                paragraph_text = paragraph.text.strip()

                # If empty paragraph, might indicate a line break
                if not paragraph_text:
                    # If we have accumulated text, add it as a text box and reset
                    if current_text_block:
                        paragraphs_text.append('\n'.join(current_text_block))
                        current_text_block = []
                else:
                    current_text_block.append(paragraph_text)

            # Add any remaining text
            if current_text_block:
                paragraphs_text.append('\n'.join(current_text_block))

            # Add all text blocks from this shape
            slide_text_boxes.extend(paragraphs_text)

        all_text_boxes.append({
            'slide_index': slide_index + 1,
            'text_boxes': slide_text_boxes
        })

    return all_text_boxes


# Create a DataFrame to visualize the modified text
def create_df_to_visualize_prs(prs):
    text_df = pd.DataFrame([
        {'Slide': item['slide_index'], 'Text Box': i + 1, 'Content': content}
        for item in extract_formatted_text(prs)
        for i, content in enumerate(item['text_boxes'])
    ])


def save_prs(path, prs):
    if type(prs) is pptx.presentation.Presentation:
        prs.save(path)
