import os
import pandas as pd
import pptx
import re
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
import io
import tempfile

# Define the scopes for Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

def authenticate_google_drive():
    """
    Authenticate with Google Drive API
    
    Returns:
        Google Drive API service
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no valid credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    # Build the Drive API client
    service = build('drive', 'v3', credentials=creds)
    return service

def search_text_in_pptx(file_path, search_text):
    """
    Search for text in all slide titles of a PowerPoint file
    
    Args:
        file_path: Path to the PowerPoint file
        search_text: Text to search for in slide titles
        
    Returns:
        bool: True if the text is found in any slide title, False otherwise
    """
    try:
        presentation = pptx.Presentation(file_path)
        if not presentation.slides:
            return False
        
        # Check all slides for the search text in titles
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame and hasattr(shape, 'is_title') and shape.is_title:
                    title_text = shape.text_frame.text.strip()
                    if search_text.lower() in title_text.lower():
                        return True
                        
        return False
    except Exception as e:
        print(f"Error searching text in {file_path}: {str(e)}")
        return False

def download_and_search_file(service, file_id, search_text):
    """
    Download a file from Google Drive and search for text in slide titles
    
    Args:
        service: Google Drive API service
        file_id: ID of the file to download
        search_text: Text to search for in slide titles
        
    Returns:
        tuple: (file_name, bool) - File name and whether it contains the search text
    """
    try:
        # Get file metadata
        file_metadata = service.files().get(fileId=file_id).execute()
        file_name = file_metadata.get('name', 'Unknown')
        
        # Create request to download the file
        request = service.files().get_media(fileId=file_id)
        
        # Download the file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_file:
            downloader = MediaIoBaseDownload(temp_file, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            
            temp_path = temp_file.name
        
        # Search for text in the downloaded file
        contains_text = search_text_in_pptx(temp_path, search_text)
        
        # Clean up the temporary file
        os.unlink(temp_path)
        
        return file_name, contains_text
    except Exception as e:
        print(f"Error processing file {file_id}: {str(e)}")
        return "Unknown", False

def find_matching_presentations(search_text, local_path=None):
    """
    Find presentations in Google Drive that have a slide title containing the search text
    
    Args:
        search_text: Text to search for in slide titles
        local_path: Optional path to search for local files
        
    Returns:
        list: List of matching file names
    """
    matching_files = []
    
    # Search Google Drive
    try:
        service = authenticate_google_drive()
        
        # Search for PowerPoint and Google Slides files
        query = "mimeType='application/vnd.google-apps.presentation' or mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation'"
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)'
        ).execute()
        
        files = results.get('files', [])
        
        for file in files:
            file_name, matches = download_and_search_file(service, file['id'], search_text)
            if matches:
                matching_files.append(file_name)
    
    except Exception as e:
        print(f"Error searching Google Drive: {str(e)}")
    
    # Search local directory if provided
    if local_path:
        for root, _, files in os.walk(local_path):
            for file in files:
                if file.lower().endswith('.pptx'):
                    file_path = os.path.join(root, file)
                    if search_text_in_pptx(file_path, search_text):
                        matching_files.append(file)
    
    return matching_files

def search_presentations(search_text='gong', search_local_path=None):
    """
    Search for presentations containing the specified text in slide titles
    
    Args:
        search_text: Text to search for in slide titles
        search_local_path: Optional path to search for local files
        
    Returns:
        list: List of matching file names
    """
    if not search_text:
        print("No search text provided")
        return []
    
    print(f"Searching for presentations with titles containing: {search_text}")
    matching_files = find_matching_presentations(search_text, search_local_path)
    
    return matching_files

# Example usage
if __name__ == "__main__":
    # Replace with your search text
    search_text = "Project Overview"
    matches = search_presentations(search_text, "path/to/local/files")
    
    print(f"Searching for: {search_text}")
    print(f"Found {len(matches)} matching presentations:")
    for match in matches:
        print(f"- {match}")
