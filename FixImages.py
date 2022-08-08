#!/usr/local/bin/python
import os, fnmatch, subprocess, shutil, zipfile
from zipfile import ZipFile

import xml.etree.ElementTree as xml_ET # for parsing xml tree
import xml.etree.ElementTree as svg_ET # for parsing svg tree

import requests, http.client
from dotenv import load_dotenv

# REST API STUFF
load_dotenv()
DOCUMOTO_API_ENDPOINT_URL = "https://integration.digabit.com/api/ext/publishing/upload/v1?submitForPublishing=true" # documoto integration URL
DOCUMOTO_API_KEY = os.environ.get('DOCUMOTO_API_KEY') # API key is stored in DOCUMOTO_API_KEY env variable
DOCUMOTO_USERNAME = "integration@nikolamotor.com"

headers = {
    # 'Content-Type': 'multipart/form-data', # comment this out to let requests handle type definition
    'Accept': 'text/plain',
    'Authorization': DOCUMOTO_API_KEY # store api key in headers dictionary with key "Authorization" as required by documoto REST API
}

exclude_list = [] # initialize exclude_list as new empty list

def register_all_namespaces(filename, ET): # function for registering xml namespaces in etree (parses original file and retains namespaces)
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        ET.register_namespace(ns, namespaces[ns])
    return namespaces # return namespaces for use in tree iteration later

def printResponseDetails(response): # function to print details of http request response (takes response object as parameter)
    method_len = len(response.request.method)
    url_len = len(response.request.url)
    headers_len = len('\r\n'.join('{}{}'.format(k, v) for k, v in response.request.headers.items()))
    body_len = len(response.request.body if response.request.body else [])
    print(f'Request size {method_len + url_len + headers_len + body_len}')

    print("RESPONSE - Code " + str(response.status_code) + ": " + http.client.responses[response.status_code]) # convert response code to description and output to console
    if response.status_code == 200: print("DOCUMOTO JOB #: " + response.text)
    else: print(response.text)

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
        os.makedirs(new_file_directory) # create new directory for modified files to be stored prior to repackaging into original PLZ archives

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

    # MODIFY XML and save to new file directory (update name and description attributes in page translation)
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for xmlFileName in fnmatch.filter(files, "*.xml"): # iterate through only the file that match specified extension
            xmlFilePath = os.path.join(path, xmlFileName) # join path and filename to get absolute file path

            # get page title by parsing filename (last element of space-delimited filename prior to extension)
            pageTitleWithUnderscores = xmlFileName.replace('.xml', '') # start by removing file extension
            pageTitleWithUnderscores = pageTitleWithUnderscores.replace('  ', ' ') # replace double spaces with single spaces
            splitChar = ' ' # define the delimiter
            listOfValues = pageTitleWithUnderscores.split(splitChar) # create list of values
            if len(listOfValues) > 0: # check for zero length
                pageTitleWithUnderscores = listOfValues[-1] # get last index of list and set pageTitle equal to it
            else:
                break # break the loop if zero length list
            pageTitleWithSpaces = pageTitleWithUnderscores.replace('_', ' ') # replace underscores with spaces
            
            # use xml.etree to update page attributes 'name' and 'description'
            XMLnamespaces = register_all_namespaces(xmlFilePath, xml_ET) # register namespaces
            XMLtree = xml_ET.parse(xmlFilePath) # parse xml file into xml tree
            XMLroot = XMLtree.getroot() # get the root tree (in this case, root is <Page>)
            translationElement = XMLroot.find("Translation", XMLnamespaces) # define Translation element so we can set its attributes
            translationElement.set('name', pageTitleWithSpaces) # set name
            translationElement.set('description', pageTitleWithSpaces) # set description

            # if it does, we can safely build and add the <Attachment> element and <Comments> subelement
            # to the root tree. <Comments> contains the filename.  We will need to upload this file also
            # build and add <Attachment> element here:
            thumbnailFileName = xmlFileName.replace('.xml', '.png') # get thumbnail filename using xml filename by replacing xml with png
            thumbnailFilePath = os.path.join(os.path.abspath(directory), thumbnailFileName) # join root directory path with thumbnail filename to get thumbnail absolute path
            if os.path.exists(thumbnailFilePath): # check to see if PNG thumbnail exists
                print(f"MATCHING THUMBNAIL FOUND: {thumbnailFilePath}")
                print("ADDING THUMBNAIL DATA TO XML")

                # add subelements
                attachmentElement = xml_ET.SubElement(XMLroot, "Attachment") # add <Attachment> subelement to <Page> root
                commentsElement = xml_ET.SubElement(attachmentElement, "Comments") # add <Comments> subelement to <Attachment> element

                # set attributes for newly created subelements
                attachmentElement.set('fileName', thumbnailFileName) # set fileName attribute
                attachmentElement.set('global', "false") # set global attribute
                attachmentElement.set('publicBelowOrg', "false") # set publicBelowOrg attribute
                attachmentElement.set('type', "THUMBNAIL") # set type attribute
                attachmentElement.set('userName', DOCUMOTO_USERNAME) # set userName attribute
                commentsElement.text = thumbnailFileName # set comments text to thumbnail filename
                
            # save the modified xml in new file directory
            newXmlFilePath = os.path.join(new_file_directory, xmlFileName)
            XMLtree.write(newXmlFilePath, encoding='utf-8', xml_declaration=True) # need to set xml_declaration to True to preserve first line of xml <?xml version="1.0" encoding="UTF-8"?>

            print(f'saved new xml to:\n{newXmlFilePath}\n')
            
        print("PAGE NAME AND DESCRIPTION MODIFICATION COMPLETE!")
        print("")
        break # prevent descending into subfolders

    # now that all files are extracted, use etree to modify attributes in the svg files to prepare them for overlaying onto rasters
    # (these modified SVGs will NOT be repacked into original PLZ archives - they're only for temporary use)
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for svgFileName in fnmatch.filter(files, "*.svg"): # iterate through only the file that match specified extension
            svgFilePath = os.path.join(path, svgFileName) # join path and filename to get absolute file path

            SVGnamespaces = register_all_namespaces(svgFilePath, svg_ET) # register namespaces
            SVGtree = svg_ET.parse(svgFilePath) # get tree
            SVGroot = SVGtree.getroot() # get root

            for text in SVGroot.iter('{http://www.w3.org/2000/svg}text'): # iterate through root to find <text> elements
                text.set('font-size', '36px') # add font-size attribute to increase text size
                text.set('dy', '9px') # add y-offset attribute to center text in bubble

            for ellipse in SVGroot.iter('{http://www.w3.org/2000/svg}ellipse'): # iterate through root to find <ellipse> elements
                ellipse.attrib.pop('stroke') # remove stroke attribute and stroke-width attributes to
                ellipse.attrib.pop('stroke-width') # remove the white outline on callout bubbles

            SVGtree.write(svgFilePath) # overwrite SVG with new file

            print(f"completed attribute modification in {svgFileName}...")

        print("")        
        print("SVG ATTRIBUTE MODIFICATION IS COMPLETE!")
        print("")
        break # prevent descending into subfolders

    # now we need to mogrify
    process = subprocess.Popen(f'mogrify -path {new_file_directory} -format png {temp_directory}/*.svg', # run the shell command, performing mogrify on ALL SVGs in the temp_directory with their corresponding PNG raster images
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
        
        print("")
        print("SOURCE PNG and XML REMOVAL FROM ORIGINAL PLZ ARCHIVES IS COMPLETE!")
        print("")
        break # prevent descending into subfolders


    # now we need to re-pack the new png and xml files into their original archives
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file

            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            pngFileName = plzFileName.replace('.plz', '.png') # raster png in the plz package (NOT the thumbnail)
            pngFilePath = os.path.join(new_file_directory, pngFileName) # join new_file_directory with pngFileName to get png absolute path
            xmlFileName = plzFileName.replace('.plz', '.xml') # find and replace .plz with .xml to get the xml file name
            xmlFilePath = os.path.join(new_file_directory, xmlFileName) # join new_file_directory with xmlFileName to get xml absolute path
            
            thumbnailFilePath = plzFilePath.replace('.plz', '.png') # this is the thumbnail, located in same root directory as PLZs
            thumbnailFileName = pngFileName # thumbnail filename should be identical to png file name (even though they're different files)

            print("")
            print(f"re-packing new PNG and XML into archive {plzFileName}")
            print("ARCHIVE CONTENTS:")
            with ZipFile(plzFilePath, 'a') as archive: # append
                archive.write(pngFilePath, arcname=pngFileName) # write the new png file to the archive
                archive.write(xmlFilePath, arcname=xmlFileName) # write the new xml file to the archive
                archive.printdir()
                
            # publish to tenant?
            if publishToDocumoto:
                # first, attempt to upload the PLZ's corresponding thumbnail image
                if os.path.exists(thumbnailFilePath): # check to see if the thumbnail even exists:
                    print("")
                    print(f"UPLOADING THUMBNAIL IMAGE: {thumbnailFileName}...")

                    filesToUpload = {'file': (thumbnailFileName, open(thumbnailFilePath, 'rb'), 'application/octet-stream')} # use filename as first parameter of 3-tuple

                    thumbnailResponse = requests.request('POST', DOCUMOTO_API_ENDPOINT_URL, headers = headers, files = filesToUpload)
                    printResponseDetails(thumbnailResponse) # print response details

                # now, upload the PLZ
                print("")
                print(f"UPLOADING AND PUBLISHING PLZ: {plzFileName}...")

                filesToUpload = {'file': (plzFileName, open(plzFilePath, 'rb'), 'application/octet-stream'), # use filename as first parameter of 3-tuple
                                 'submitForPublishing': True} # set submitForPublishing to true so file is published once uploaded

                plzResponse = requests.request('POST', DOCUMOTO_API_ENDPOINT_URL, headers = headers, files = filesToUpload)
                printResponseDetails(plzResponse) # print response details

        print("")
        print("PNG AND XML REPACKING INTO ORIGINAL PLZ ARCHIVES IS COMPLETE!")
        if publishToDocumoto: print("ARCHIVES HAVE BEEN UPLOADED AND PUBLISHED TO THE DOCUMOTO TENANT!")
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
print("************************")
print("If you wish to publish THUMBNAIL images as well, they need to be located")
print("in the same root directory as the PLZ archives, for XML modification and")
print("automatic uploading of thumbnails to occur.")
print("************************")
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

# check directory contents, warn user if no thumbnails are found (they can still be moved to the input directory by the user at this point if necessary)
thumbnailsMissing = False # initialize this flag to false (we'll set to true if we find at least one PLZ archive without a corresponding thumbnail image)
print(f"Relevant files found in {directory}:")
for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
    # list PLZ files in directory
    print("***PLZ FILES: ***")
    for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the files that match plz file extension
        plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
        thumbnailFilePath = plzFilePath.replace('.plz', '.png') # this is the thumbnail, located in same root directory as PLZs
        print(f"\t{plzFileName}") # print the plz filename
        if not os.path.exists(thumbnailFilePath):
            thumbnailsMissing = True
            print("*****WARNING: THUMBNAIL IMAGE FOR THE ABOVE PLZ ARCHIVE IS MISSING!*****")

    print("")
    
    # list png thumbnail files in directory
    print("***PNG THUMBNAIL FILES: ***")
    for thumbnailFileName in fnmatch.filter(files, "*.png"): # iterate through only the files that match png file extension
        print(f"\t{thumbnailFileName}") # print the png thumbnail filename
    break # prevent descending into subfolders

print("")
if thumbnailsMissing:
    print("*********************************************************************************************************")
    print("WARNING: AT LEAST 1 PLZ ARCHIVE IS MISSING ITS CORRESPONDING THUMBNAIL IMAGE IN THE ROOT DIRECTORY!!!")
    print(f"You can still copy the thumbnail(s) into {directory}")
    print("RIGHT NOW if you'd like to include them. Otherwise, they'll need to be added manually afterwards within the tenant.")
    print("*********************************************************************************************************")
    print("")
    


publishToDocumoto = input("Publish the new PLZ pages to Documoto? Type YES to publish: ").upper() == "YES"
print("")
cleanup = input("Cleanup temporary files after repackaging archives? Type YES to delete: ").upper() == "YES"
print("")


fixTheFiles(directory) # FIX THE FILES!

print("============================================")
print("                 GOODBYE!")
print("============================================")

