# -*- coding: utf-8 -*-
"""
Extracting data from OutLook emails according to provided query
uses methods here:
    https://msdn.microsoft.com/en-us/library/office/dn467914.aspx
    https://msdn.microsoft.com/en-us/library/ms527464(v=exchg.10).aspx
    https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_properties.aspx    

@author: V.Nesterov

"""
from os import sys
import os
import time
import win32com.client
import zipfile
import re
from tqdm import tqdm

class Mailbox():
    '''
        #######################################################################    
        Extracts data from specified mailbox according to predefined criteria.
        #######################################################################

        Takes following arguments:
            - path - as string: full path to directory to download content into (default = None)
            - mailbox - as string: specific text contained in mailbox name to identify and open
            - folder - as integer: 0 deleted items; 1 for inbox; 2 for outbox; 3 for sent (index may differ depending on structure of folders)
            - subj_keys - as list of strings: list of strings to search in email subject line
                (exact match only, not case sensitive)
            - text - boolean: to scrap text body of message and save as .txt file (default = False)
            - attach - boolean: to upload attachment (default = False)
            - unzip - boolean: to unzip attachment and remove zip file if .zip archive attached (default = False)
    '''
    def __init__(self, path=None, mailbox=None, folder=1, subj_keys=None, text=False, attach=False, unzip=False):
        if path[-1] != '/':
            self.path = path+'/'
        else:
            self.path = path
        self.mailbox = mailbox
        self.folder = folder
        self.subj_keys = subj_keys
        self.text = text
        self.attach = attach
        self.unzip = unzip
        
        if path == None:
            print("Please specify path to upload attachments. \nTerminating... \n")
            sys.exit()
        if mailbox==None:
            print("Please define specific text contained in mailbox name to identify and open. \n", help(Mailbox))
            sys.exit()
        if subj_keys==None:
            print("Please specify subj_keys - as list of strings: "+\
                  "list of strings to search in email subject line. \n", help(Mailbox)) 
            sys.exit()
        if text==False and attach==False:
            print("Please specify option you want to use with this tool. \n", help(Mailbox))
            sys.exit()
                  
    def unzip_file(self, locale, filename):
        '''unzipping function'''
        try:
            with zipfile.ZipFile(filename, 'r') as zip_ref:
                # extracting
                print("Unzipping:", filename)
                zip_ref.extractall(locale+"/")
                zip_ref.close()
                # deleting zip archive
                print("Deleting zip file", filename)
                os.remove(filename)
        except:
            print("\n! Error unzipping", filename) 
        
    def search_mail(self):
        ''' Iterates through each email in mailbox applying specified search criteria
        '''
        print("\nRetrieving data")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # accessing mailbox folders and parsing
        bx = None
        print("\nLooking for '{}' in mailboxes...".format(self.mailbox))
        for b in tqdm(range(100)): # trying to pick mailbox number
            try:
                outlook_folder = outlook.Folders[b].Folders[self.folder]
                bx = str(outlook.Folders[b]).split(',')[0]
                if self.mailbox in bx:
                    print("\nFound mailbox {}.".format(bx))
                    break
                else:
                    bx = None
            except:
                pass
        if bx == None:
            print("ERROR: No '{m}' key found in mailboxes. Check if Outlook has mailboxes with '{m}' key in name.".format(m=self.mailbox))
            print("Terminating...")
            sys.exit()
        
        print("\nOpened: ", bx)
        time.sleep(2)
        print("Found:", outlook_folder)
        time.sleep(2)
    
        msg = outlook_folder.Items
        # checking each message
        for i in tqdm(msg, desc="\nLooking into emails... ", unit="email "):
            # iterating through keywords to find in subject line
            for s in self.subj_keys:
                try:
                    if s.lower() in i.Subject.lower():
                        print("\nfound key '{}' in email - parsing".format(s))
                        # defining path and name
                        msg_dir = self.path+re.sub('[^\w\-_\.s]', '_', i.Subject[:25])+"/"
                        file_name = re.sub('[^\w\-_\.s]', '_', i.Subject[:40])+".txt"
                        # if extracting text
                        if self.text:
                            print("Extracting body text")
                            #creating directory
                            try:
                                # Create target Directory
                                os.mkdir(msg_dir)
                            except:
                                pass
                            new_text_file = open(msg_dir+file_name, "w+")
                            new_text_file.write(i.Body)
                            new_text_file.close()
                        # if iterating through attachments and uploading
                        if self.attach:
                            for att in i.Attachments:
                                print("Found attachement '{}' - saving".format(att))
                                # saving attachment
                                fullname = msg_dir+str(att)
                                att.SaveAsFile(fullname)
                                # if unzipping required
                                if self.unzip and ".zip" in fullname:
                                    self.unzip_file(msg_dir, fullname)
                except:
                    print("Unexpected error reading email")
                    raise
        print("Completed searching and parsing. Check '{}' for output".format(self.path))
    