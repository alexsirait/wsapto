import os
from django.conf import settings

class SingleFileUploadHandler:
    def __init__(self, file):
        self.file = file

    def save_file(self):
        if not self.file:
            return {'success': False, 'message': 'No file was uploaded.', 'messagetype': 'E'}

        # Create directory if it doesn't exist
        upload_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
        os.makedirs(upload_dir, exist_ok=True)

        # Save the file
        file_path = os.path.join(upload_dir, self.file.name)

        try:
            with open(file_path, 'wb+') as destination:
                for chunk in self.file.chunks():
                    destination.write(chunk)
            return {'success': True, 'message': 'File uploaded successfully.', 'file': self.file.name, 'messagetype': 'S'}
        except Exception as e:
            return {'success': False, 'message': str(e), 'messagetype': 'E'}

class MultipleFileUploadHandler:
    def __init__(self, files):
        self.files = files

    def save_files(self):
        if not self.files:
            return {'success': False, 'message': 'No files were uploaded.', 'messagetype': 'E'}

        saved_files = []
        upload_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
        os.makedirs(upload_dir, exist_ok=True)

        for file in self.files:
            file_path = os.path.join(upload_dir, file.name)
            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in file.chunks():
                        destination.write(chunk)
                saved_files.append(file.name)
            except Exception as e:
                return {'success': False, 'message': f'Failed to upload {file.name}: {str(e)}', 'messagetype': 'E'}

        return {'success': True, 'message': 'Files uploaded successfully.', 'files': saved_files, 'messagetype': 'S'}
