from __future__ import print_function
import httplib2
import os
import io

from apiclient import discovery
from apiclient.http import MediaIoBaseDownload
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Form2Email'


def get_credentials(credfile):
    home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'Form2EmailGDrive.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(
            os.path.join(home_dir, credfile), SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def getDriveID(url):
    driveid = ''
    for x in range(0, len(url)):
        if url[x:x + 3] == 'id=':
            for y in range(x + 3, len(url)):
                if url[y] == '/':
                    break
                else:
                    driveid += url[y]
            break
    return driveid


def download(fileURL, credfile=CLIENT_SECRET_FILE):
    """
    Creates a Google Drive API service object and downloads file specified
    """
    credentials = get_credentials(credfile)
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)

    home_dir = os.path.abspath(os.path.join(__file__, os.pardir))
    image_dir = os.path.join(home_dir, '.images')
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)
    os.chdir(image_dir)

    file_id = getDriveID(fileURL)
    request = drive_service.files().get_media(fileId=file_id)
    filename = file_id + '.tiff'
    fh = io.FileIO(filename, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Downloading " + filename + " %d%%." %
              int(status.progress() * 100))
    return filename
