#!/usr/local/bin/python
import os, fnmatch, subprocess, shutil, zipfile
from zipfile import ZipFile

import requests, http.client
from dotenv import load_dotenv

# REST API STUFF
load_dotenv()
DOCUMOTO_API_ENDPOINT_URL = "https://integration.digabit.com/api/ext/publishing/upload/v1?submitForPublishing=true" # documoto integration URL
DOCUMOTO_API_KEY = os.environ.get('DOCUMOTO_API_KEY') # API key is stored in DOCUMOTO_API_KEY env variable

headers = {
    # 'Content-Type': 'multipart/form-data', # comment this out to let requests handle type definition
    'Accept': 'text/plain',
    'Authorization': DOCUMOTO_API_KEY # store api key in headers dictionary with key "Authorization" as required by documoto REST API
}

exclude_list = [] # initialize exclude_list as new empty list

def isExcluded(plzFileName):
    if plzFileName in exclude_list: # check exclude list for the PLZ filename
        print("")
        print(f"***** SKIPPING {plzFileName} *****")
        print("please process this file in Docustudio first!")
        print("")
        return True
    else:
        return False

def fixTheFiles(directory): # this function decompresses all PLZs in given directory into same directory
    
    temp_directory = os.path.abspath(directory + "/FixImages_temp") # define temporary directory name
    new_file_directory = os.path.abspath(temp_directory + "/new_files") # define new_file temporary directory name
    if not os.path.exists(new_file_directory):
        print("")
        print(f'creating directory {new_file_directory} for temporary storage of new files.')
        print("")
        os.makedirs(new_file_directory) # create new directory (because mogrify is dumb and can't create one itself)

    # Extract all the contents of zip file in temporary subdirectory
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path

            # need to check if a PNG file exists within the archive. if missing, PLZ has not yet been
            # processed in docustudio (in other words, callout bubbles are not yet added)
            zf = zipfile.ZipFile(plzFilePath, 'r')
            processedInDocustudio = not any("_page_documoto" in item for item in zf.namelist()) # check to see if "_page_documoto" appears anywhere in the file name list
            if not processedInDocustudio:
                print("")
                print("****************************************")
                print(f"WARNING! {plzFileName} HAS NOT YET BEEN PROCESSED IN DOCUSTUDIO!")
                print("PLEASE PROCESS IN DOCUSTUDIO, ADDING CALLOUT BUBBLES FIRST")
                print(f"{plzFileName} has been skipped!")
                print("****************************************")
                print("")
                exclude_list.append(plzFileName) # add unprocessed PLZ to exclude list
                continue # continue to next iteration of this loop, skipping this file
            zf.close()

            with ZipFile(plzFilePath, 'r') as archive:
                archive.extractall(temp_directory)
                print(f"extracting {plzFileName}...")
        break # prevent descending into subfolders
    print(f"PLZ archives have been extracted into {temp_directory}!")
    print("")

    # MODIFY XML and move to new file directory
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for xmlFileName in fnmatch.filter(files, "*.xml"): # iterate through only the file that match specified extension
            xmlFilePath = os.path.join(path, xmlFileName) # join path and filename to get absolute file path

            # get page title from filename (last element of space-delimited filename prior to extension)
            pageTitleWithUnderscores = xmlFileName.replace('.xml', '') # start by removing file extension
            pageTitleWithUnderscores = pageTitleWithUnderscores.replace('  ', ' ') # replace double spaces with single spaces
            splitChar = ' ' # define the delimiter
            listOfValues = pageTitleWithUnderscores.split(splitChar) # create list of values
            if len(listOfValues) > 0: # check for zero length
                pageTitleWithUnderscores = listOfValues[-1] # get last index of list and set pageTitle equal to it
            else:
                break # break the loop if zero length list

            # build string we want to replace in the xml
            oldNameAttribute = xmlFileName.replace(' ' + pageTitleWithUnderscores + '.xml', '') # remove final space, page title, and file extension from the xml file name to get the old name attribute value
            stringToFind = '<Translation locale=\"en_US\" name=\"' + oldNameAttribute + '\" description=\"\"/>'

            # build replacement string
            pageTitleWithSpaces = pageTitleWithUnderscores.replace('_', ' ') # replace underscores with spaces
            replacementString = '<Translation locale=\"en_US\" name=\"' + pageTitleWithSpaces + '\" description=\"' + pageTitleWithSpaces + '\"/>'
            
            with open(xmlFilePath) as f:
                s = f.read() # open the file
            s = s.replace(stringToFind, replacementString) # find and replace
            print("")
            print(f'in {xmlFileName}, replacing:\n{stringToFind}\nwith:\n{replacementString}') # notify user
            print("")

            with open(xmlFilePath, "w") as f:
                f.write(s) # close the file

            # move modified xml to new file directory
            newXmlFilePath = os.path.join(new_file_directory, xmlFileName)
            shutil.move(xmlFilePath, newXmlFilePath)
            print(f'moving\n{xmlFilePath}\nto\n{newXmlFilePath}')
            
        print("find-and-replace complete!")
        print("")
        break # prevent descending into subfolders

    # now that all files are extracted, let's find and replace in the svg files to prepare them for overlaying onto rasters (these modified SVGs will eventually be discarded)
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for svgFileName in fnmatch.filter(files, "*.svg"): # iterate through only the file that match specified extension
            svgFilePath = os.path.join(path, svgFileName) # join path and filename to get absolute file path
            with open(svgFilePath) as f:
                s = f.read() # open the file
            s = s.replace('stroke=\"#FFFFFF\" stroke-width=\"2\" ', '') # find and replace (callout bubble white outline)
            s = s.replace('<text', '<text font-size=\"36px\" dy=\"9px\"') # find and replace (callout text size and alignment)
            with open(svgFilePath, "w") as f:
                f.write(s) # close the file
            print(f"performing find-and-replace on {svgFileName}...")
        
        print("find-and-replace complete!")
        print("")
        break # prevent descending into subfolders

    # now we need to mogrify
    process = subprocess.Popen(f'mogrify -path {new_file_directory} -format png {temp_directory}/*.svg', # run the shell command
                           shell=True, stdout=subprocess.PIPE)
    process.wait() # wait for process to finish in current thread before proceeding
    print(f"newly-generated PNGs are located in {new_file_directory}!")
    print("")

    # we need to remove the original png and xml files from the plz archives to avoid duplication
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file

            old_filepath = os.path.join(path, plzFileName)
            new_filepath = os.path.join(path, plzFileName + "_new")

            zin = zipfile.ZipFile(old_filepath, 'r') # old file
            zout = zipfile.ZipFile(new_filepath, 'w') # new file
            for item in zin.infolist():
                buffer = zin.read(item.filename)
                if ((item.filename[-4:] != '.png') and (item.filename[-4:] != '.xml')): # write all except png amd xml files to new archive
                    zout.writestr(item, buffer)
            zout.close() # close files
            zin.close() # close files
            os.remove(old_filepath) # delete old file
            os.rename(new_filepath, old_filepath) # rename new file to old filename
            print(f"removing source png and xml files from original archive {plzFileName}")
        
        print("source PNG and XML removal from original PLZ archives is complete!")
        print("")
        break # prevent descending into subfolders


    # now we need to re-pack the new png and xml files into their original archives
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file

            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            pngFileName = plzFileName.replace('.plz', '.png')
            pngFilePath = os.path.join(new_file_directory, pngFileName) # remove last 3 characters from plz file path (.plz) to get base filename (no extension).
                                                                       # Then, append "png" to it to get the png file path (after joining with new_file_directory path)
            xmlFileName = plzFileName.replace('.plz', '.xml')
            xmlFilePath = os.path.join(new_file_directory, xmlFileName) # remove last 3 characters from plz file path (.plz) to get base filename (no extension).
                                                                       # Then, append "xml" to it to get the xml file path (after joining with new_file_directory path)
            print("")
            print(f"re-packing new PNG and XML into archive {plzFileName}")
            print("ARCHIVE CONTENTS:")
            with ZipFile(plzFilePath, 'a') as archive: # append
                archive.write(pngFilePath, arcname=pngFileName) # write the new png file to the archive
                archive.write(xmlFilePath, arcname=xmlFileName) # write the new xml file to the archive
                archive.printdir()
                
            # publish to tenant?
            if publishToDocumoto:
                print("PUBLISHING...")

                filesToUpload = {'file': (plzFileName, open(plzFilePath, 'rb'), 'application/octet-stream')} # use filename as first parameter of 3-tuple

                response = requests.request('POST', DOCUMOTO_API_ENDPOINT_URL, headers = headers, files = filesToUpload)

                # print request size
                method_len = len(response.request.method)
                url_len = len(response.request.url)
                headers_len = len('\r\n'.join('{}{}'.format(k, v) for k, v in response.request.headers.items()))
                body_len = len(response.request.body if response.request.body else [])
                print(f'Request size {method_len + url_len + headers_len + body_len}')

                print("RESPONSE - Code " + str(response.status_code) + ": " + http.client.responses[response.status_code]) # convert response code to description and output to console
                print(response.text)

        print("")
        print("PNG and XML re-packing is completed!")
        print("")
        break # prevent descending into subfolders

    # clean up (delete temporary directory)
    if cleanup:
        shutil.rmtree(temp_directory)
        print(f"deleting temp directory {temp_directory}")
        print("")
    else:
        print(f"keeping temporary directory! files are located in {temp_directory}")
        print("")


    print("")
    print("PROCESS IS COMPLETE!!!")
    print("")

    # report exlusion list
    if len(exclude_list) != 0:
        print("***** WARNING *****")
        print("the following PLZ files have been skipped:")
        print("")
        for item in exclude_list:
            print(item)
        print("")
        print("PLEASE PROCESS THESE IN DOCUSTUDIO FIRST, ADDING CALLOUT BUBBLES, AND TRY AGAIN!")
        print("")


print("")
print("")
print("")
print("Welcome to Documoto Image Fixer!")
print("================================")
print("This is an application to embed Documoto callout bubbles onto the source raster images.")
print("The directory entered below must include documoto package files (plz).")
print("Files will be extracted, modifications made, and repackaged into the original archives.")
print("")

# preferences
isValidDirectory = False
while not isValidDirectory:
    directory = input("Enter full directory containing the plz files: ")
    directory = directory.replace('\\', '/') # replace backslashes with forward slashes
    print("")
    if os.path.exists(directory): # check to see if directory exists before proceeding
        if not ' ' in directory: # check to see if directory has spaces
            isValidDirectory = True # if no spaces, flag this to true to prevent next loop iteration
        else: # complain
            print("")
            print(f"{directory} contains spaces. Please use a directory structure containing no spaces.")
            print("Please try again!")
            print("")
    else: # complain
        print("")
        print(f"{directory} is not a valid path.")
        print("Please try again!")
        print("")
publishToDocumoto = input("Publish the new PLZ pages to Documoto? Type YES to publish: ").upper() == "YES"
print("")
cleanup = input("Cleanup temporary files after repackaging archives? Type YES to delete: ").upper() == "YES"
print("")


fixTheFiles(directory) # FIX THE FILES!

print("============================================")
print("                 GOODBYE!")
print("============================================")

