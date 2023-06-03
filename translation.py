import re
from googletrans import Translator
from string import punctuation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
from pptx import Presentation
import win32com.client as client
import win32com.client.gencache as gencache
from docx import Document
import urllib.parse
import io
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from googleapiclient.discovery import build

start_translating = False

# Initialize translator
translator = Translator()

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'googleKey.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

def remove_illegal_chars(filename):
    illegal_chars = r"[']"  # Define the pattern for illegal characters (single quote in this case)
    return re.sub(illegal_chars, "", filename)  # Use regular expression to remove illegal characters

def is_punctuation(text):
    return all(char in punctuation for char in text)

def translate_to_english(text):
    if len(text.strip()) < 2 or is_punctuation(text.strip()):
        return text  # Return the original text if it has less than 2 non-whitespace characters or consists only of punctuation

    try:
        translation = translator.translate(text, src="zh-CN", dest="en")  # Translate the text from Chinese to English
        translated_text = translation.text
        #translated_text = translated_text.replace("'", "")  # remove apostrophes
        return translated_text
    except Exception as e:
        print(f"Failed to translate text: {text}. Error: {str(e)}")
        return text  # Return the original text if translation fails

def translate_text_frame(text_frame):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if not run.text:
                continue  # Skip empty runs
            try:
                translated = translate_to_english(run.text)  # Translate the text to English
                run.text = translated  # Update the text with the translated version
            except Exception as e:
                print(f"Failed to translate text: {run.text}. Error: {str(e)}")
                continue

def adjust_text_size(shape):
    while shape.text_frame.text != "":
        try:
            # Try to access the last character of the shape's text
            _ = shape.text_frame.text[-1]
            break  # Break the loop if the last character can be accessed
        except IndexError:
            # If the last character cannot be accessed, the text overflows the shape
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # Decrease the font size by 1 point
                    run.font.size = Pt(run.font.size.pt - 1)

def process_shape(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        # If the shape is a group, recursively process each shape within the group
        for shape_in_group in shape.shapes:
            process_shape(shape_in_group)
    elif shape.has_text_frame:
        # If the shape has a text frame, translate the text and adjust the text size
        translate_text_frame(shape.text_frame)
        adjust_text_size(shape)
    elif shape.has_table:
        # If the shape is a table, process each cell's text frame and adjust text size
        for row in shape.table.rows:
            for cell in row.cells:
                translate_text_frame(cell.text_frame)
                adjust_text_size(cell)

def translate_pptx(pptx_path):
    pres = Presentation(pptx_path)

    # Process each slide in the presentation
    for slide in pres.slides:
        for shape in slide.shapes:
            process_shape(shape)
            if shape.has_text_frame:
                translate_text_frame(shape.text_frame)
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        translate_text_frame(cell.text_frame)

    directory, filename = os.path.split(pptx_path)
    filename_without_ext = os.path.splitext(filename)[0]

    try:
        translated_filename_without_ext = translate_to_english(filename_without_ext)
    except Exception as e:
        print(f"Failed to translate filename: {filename_without_ext}. Error: {str(e)}")
        translated_filename_without_ext = filename_without_ext

    translated_filename = translated_filename_without_ext + ".pptx"
    translated_pptx_path = os.path.join(directory, translated_filename)
    pres.save(translated_pptx_path)

    print(f"Translated presentation saved at: {translated_pptx_path}")
    return translated_pptx_path

def convert_doc_to_docx(doc_path):
    doc_path = os.path.abspath(doc_path)  # Use absolute path

    try:
        word = gencache.EnsureDispatch('Word.Application')  # Create an instance of Microsoft Word
        doc = word.Documents.Open(doc_path)  # Open the .doc file
        doc.Activate()

        new_file_abs = os.path.abspath(doc_path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)  # Rename the file with .docx extension

        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=client.wdFormatXMLDocument
        )  # Save the document in .docx format
        doc.Close(False)
        word.Quit()
    except Exception as e:
        print(f"Failed to convert file: {doc_path}. Error: {str(e)}")
        return doc_path

    return new_file_abs

def translate_docx(doc_path):
    doc = Document(doc_path)

    # Translate text in paragraphs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if not run.text:
                continue  # Skip empty runs
            try:
                translated = translate_to_english(run.text)  # Translate the text to English
                run.text = translated  # Update the text with the translated version
            except Exception as e:
                print(f"Failed to translate text: {run.text}. Error: {str(e)}")
                continue

    # Translate text in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if not run.text:
                            continue  # Skip empty runs
                        try:
                            translated = translate_to_english(run.text)  # Translate the text to English
                            run.text = translated  # Update the text with the translated version
                        except Exception as e:
                            print(f"Failed to translate text: {run.text}. Error: {str(e)}")
                            continue

    directory, filename = os.path.split(doc_path)
    filename_without_ext = os.path.splitext(filename)[0]

    try:
        translated_filename_without_ext = translate_to_english(filename_without_ext)
    except Exception as e:
        print(f"Failed to translate filename: {filename_without_ext}. Error: {str(e)}")
        translated_filename_without_ext = filename_without_ext

    translated_filename = translated_filename_without_ext + ".docx"
    translated_doc_path = os.path.join(directory, translated_filename)
    doc.save(translated_doc_path)

    print(f"Translated document saved at: {translated_doc_path}")
    return translated_doc_path

def create_folder(name, parent_id):
    name = remove_illegal_chars(name)  # Remove illegal characters from the folder name

    # Check if the folder already exists
    response = drive_service.files().list(
        q=f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder'",
        spaces='drive',
        fields='files(id, name)').execute()
    
    if response.get('files'):
        folder_id = response.get('files')[0].get('id')
        print(f"Folder '{name}' already exists in parent folder ID '{parent_id}'")
        return folder_id, False  # Return False indicating the folder already exists

    # Create the folder if it doesn't exist
    file_metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    file = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
    
    print(f"Folder '{name}' created in parent folder ID '{parent_id}'")
    return file.get('id'), True  # Return True indicating the folder is newly created

def delete_file(file_id):
    try:
        drive_service.files().delete(fileId=file_id).execute()
        print(f"File with ID {file_id} has been deleted.")
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def download_file(file_id, directory_path, file_name, translated_root_id, override=False):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)  # Create the directory if it doesn't exist

    translated_file_name = translate_file_name(file_name)
    file_path = os.path.join(directory_path, translated_file_name)

    # Check if the file already exists in the translated_root_id directory on Google Drive
    encoded_filename = urllib.parse.quote(translated_file_name)
    response = drive_service.files().list(
        q=f"name='{encoded_filename}' and '{translated_root_id}' in parents",
        fields='files(id, name)').execute()
    
    if response.get('files') and not override:
        print(f"File '{translated_file_name}' already exists in parent folder ID '{translated_root_id}'")
        return None  # Return None if the file already exists
    
    elif response.get('files') and override:
        # Get the ID of the file to delete
        file_to_delete_id = response.get('files')[0]['id']
        delete_file(file_to_delete_id)  # Delete the existing file
    
    # If the file does not exist, download it
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    with open(file_path, 'wb') as f:
        f.write(fh.getbuffer())

    return file_path

def upload_file(file_path, parent_folder_id, override=True):
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [parent_folder_id]
    }

    # Check if the file already exists in the parent_folder_id directory on Google Drive
    response = drive_service.files().list(
        q=f"name='{file_metadata['name']}' and '{parent_folder_id}' in parents",
        fields='files(id, name)').execute()
    
    if response.get('files') and not override:
        print(f"File '{file_metadata['name']}' already exists in parent folder ID '{parent_folder_id}'")
        return  # Do not upload the file if it already exists
    
    elif response.get('files') and override:
        # Get the ID of the file to delete
        file_to_delete_id = response.get('files')[0]['id']
        delete_file(file_to_delete_id)  # Delete the existing file
    
    media = MediaFileUpload(file_path)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"File uploaded with ID: {file.get('id')}")

def translate_file_name(file_name):
    file_name_without_ext, ext = os.path.splitext(file_name)

    try:
        translated_file_name_without_ext = translate_to_english(file_name_without_ext)
    except Exception as e:
        print(f"Failed to translate filename: {file_name_without_ext}. Error: {str(e)}")
        translated_file_name_without_ext = file_name_without_ext

    translated_file_name_without_ext = remove_illegal_chars(translated_file_name_without_ext)
    translated_file_name = translated_file_name_without_ext + ext
    return translated_file_name

def copy_and_rename_file(file_id, translated_root_id, translated_file_name, override=False):
    # Make a copy of the original file in the new directory
    file_metadata = {
        'name': translated_file_name,
        'parents': [translated_root_id]
    }

    # Check if the file already exists in the translated_root_id directory on Google Drive
    response = drive_service.files().list(
        q=f"name='{translated_file_name}' and '{translated_root_id}' in parents",
        fields='files(id, name)').execute()
    
    if response.get('files') and not override:
        print(f"File '{translated_file_name}' already exists in parent folder ID '{translated_root_id}'")
        return  # Do not copy the file if it already exists
    
    elif response.get('files') and override:
        # Get the ID of the file to delete
        file_to_delete_id = response.get('files')[0]['id']
        delete_file(file_to_delete_id)  # Delete the existing file
    
    copied_file = drive_service.files().copy(
        fileId=file_id,
        body=file_metadata,
        fields='id'
    ).execute()
    print(f"File copied and renamed with ID: {copied_file.get('id')}")

def process_directory(directory_id, translated_root_id, start_file=None, convert_docs=False, override_docs=False, convert_slides=False, override_slides=False, copy_translate_others=False, override_others=False):
    results = drive_service.files().list(
        q=f"'{directory_id}' in parents and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)").execute()

    items = results.get('files', [])

    for item in items:
        subdirectory_id = item['id']
        file_name = translate_to_english(item['name'])

        # If start_translating is False and the current item's name matches start_file, set start_translating to True
        global start_translating
        if start_file and not start_translating and file_name == start_file:
            start_translating = True

        translated_subdirectory_id, is_new_folder = create_folder(file_name, translated_root_id)
        if is_new_folder:  # Only process the subdirectory if it is newly created
            print(f"Processing subdirectory '{item['name']}' with ID '{subdirectory_id}'")
            process_directory(subdirectory_id, translated_subdirectory_id, start_file, convert_docs, override_docs, convert_slides, override_slides, copy_translate_others, override_others)
        else:
            print(f"Processing subdirectory '{item['name']}' with ID '{subdirectory_id}'")
            process_directory(subdirectory_id, translated_subdirectory_id, start_file, convert_docs, override_docs, convert_slides, override_slides, copy_translate_others, override_others)

    results = drive_service.files().list(
        q=f"'{directory_id}' in parents and mimeType!='application/vnd.google-apps.folder'",
        fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    local_directory_path = os.path.join('./temp_drive_files', directory_id)
    for item in items:
        # Skip processing this item if start_translating is False
        if not start_translating:
            continue

        file_name = item['name']
        file_id = item['id']

        # Query if the translated file already exists in the translated directory
        translated_file_name = translate_file_name(file_name)

        # URL encode the filename
        encoded_filename = urllib.parse.quote(translated_file_name)
        response = drive_service.files().list(
            q=f"name='{encoded_filename}' and '{translated_root_id}' in parents",
            fields='files(id, name)').execute()
        if response.get('files'):
            print(f"File '{translated_file_name}' already exists in parent folder ID '{translated_root_id}'")
            continue  # Skip this file if the translated version already exists

        # Now download and process the file only if the translated version doesn't exist
        if item['mimeType'] in ['application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'] and convert_docs:
            file_path = download_file(file_id, local_directory_path, file_name, translated_root_id, False)
            if item['mimeType'] == 'application/msword':  # If it is a .doc file
                file_path = convert_doc_to_docx(file_path)
            translated_file_path = translate_docx(file_path)
            upload_file(translated_file_path, translated_root_id, override_docs)
            os.remove(translated_file_path)  # Delete the file after uploading
        elif item['mimeType'] == 'application/vnd.openxmlformats-officedocument.presentationml.presentation' and convert_slides:
            file_path = download_file(file_id, local_directory_path, file_name, translated_root_id, False)
            translated_file_path = translate_pptx(file_path)
            upload_file(translated_file_path, translated_root_id, override_slides)
            os.remove(translated_file_path)  # Delete the file after uploading
        elif copy_translate_others:
            # If the file is not a Word document, copy and rename it without translating the content
            translated_file_name = translate_file_name(file_name)
            copy_and_rename_file(file_id, translated_root_id, translated_file_name, override_others)
