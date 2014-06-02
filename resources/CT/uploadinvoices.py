#!/usr/bin/python

import argparse
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from pydrive.drive import GoogleDrive
import os

os.environ['PWD'] = os.path.dirname(os.path.realpath(__file__))

parser = argparse.ArgumentParser(description='Uploads file to Google Drive')
parser.add_argument('file', help='an integer for the accumulator') 
args = parser.parse_args()

gauth = GoogleAuth()
gauth.CommandLineAuth()

drive = GoogleDrive(gauth)
file1 = drive.CreateFile()  # Create GoogleDriveFile instance with title 'Hello.txt'
file1.SetContentFile(args.file) # Set content of the file from given string
file1.Upload()


