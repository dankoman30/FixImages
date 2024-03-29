#!/usr/local/bin/python
import os, fnmatch, subprocess, shutil, zipfile, glob
from zipfile import ZipFile

import xml.etree.ElementTree as xml_ET # for parsing xml tree
import xml.etree.ElementTree as svg_ET # for parsing svg tree

import requests, http.client
from dotenv import load_dotenv

from tkinter import filedialog
from tkinter import *

# REST API STUFF
load_dotenv()
DOCUMOTO_API_UPLOAD_ENDPOINT_URL = "https://documoto.digabit.com/api/ext/publishing/upload/v1?submitForPublishing=true" # documoto production environment file upload URL
DOCUMOTO_API_KEY = os.environ.get('DOCUMOTO_API_KEY_PRODUCTION') # production environment API key value is stored in DOCUMOTO_API_KEY_PRODUCTION env variable
DOCUMOTO_USERNAME = "daniel.koman@nikolamotor.com"

text_headers = { # for text response content type
    'Accept': 'text/plain',
    'Authorization': DOCUMOTO_API_KEY # store api key in headers dictionary with key "Authorization" as required by documoto REST API
}


json_headers = { # for json response content type
    'Accept': 'application/json',
    'Authorization': DOCUMOTO_API_KEY # store api key in headers dictionary with key "Authorization" as required by documoto REST API
}

binary_headers = { # for binary response content type (file download)
    'Accept': 'application/octet-stream',
    'Authorization': DOCUMOTO_API_KEY # store api key in headers dictionary with key "Authorization" as required by documoto REST API
}

exclude_list = [] # initialize exclude_list as new empty list

# FUNCTION DEFINITIONS

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
    print("RESPONSE TEXT: " + response.text)

def isExcluded(plzFileName):
    if plzFileName in exclude_list: # check exclude list for the PLZ filename
        print("")
        print(f"***** SKIPPING {plzFileName} *****")
        print("please process this file in Docustudio first!")
        print("")
        return True
    else:
        return False

def getPageTitleStringFromFilename(filename): # given a filename, strips the extension and returns the final space-delimited value, after replacing underscores with spaces
    # get page title by parsing filename (last element of space-delimited filename prior to extension)
    filenameWithoutExtension = filename[:-4] # start by removing file extension, or the final 4 characters of the filename string
    filenameRemoveDoubleSpaces = filenameWithoutExtension.replace('  ', ' ') # replace double spaces with single spaces
    splitChar = ' ' # define the delimiter as a single space
    listOfValues = filenameRemoveDoubleSpaces.split(splitChar) # create list of values
    if len(listOfValues) > 0: # check for zero length
        pageTitleWithUnderscores = listOfValues[-1] # get last index of list and set pageTitle equal to it
    else:
        return "" # if zero length list, return empty string
    pageTitleWithSpaces = pageTitleWithUnderscores.replace('_', ' ') # replace underscores with spaces
    return pageTitleWithSpaces

def imageFixerIntro():
    print("")
    print("============================================")
    print("                 HELLO THERE!")
    print("============================================")
    print("")
    print("Welcome to Documoto Image Fixer!")
    print("================================")
    print("This application processes PLZ page files and automates several tasks, including:")
    print("     * embed callout bubbles onto the source raster images so they appear properly in PDF catalog")
    print("     * set page title in xml based on filename string")
    print("     * link thumbnail images in xml")
    print("     * publish linked thumbnails to Documoto tenant")
    print("     * link hotpoint callouts to other pages (optional)")
    print("     * publish page files to Documoto tenant (optional)")
    print("Files will be extracted, modified, and repackaged into their original archives.")
    print("")
    print("If you wish to link and publish THUMBNAIL images, they need to be located")
    print("in the same root directory as the PLZ archives, and have the same base filenames as")
    print("their corresponding pages.")
    print("")

def getRootDirectoryFromUser():
    print("")
    print("")
    print("")
    print("")
    print("")
    print("*******************************************************************************")
    print("* PLEASE SELECT THE DIRECTORY CONTAINING THE PLZ FILES YOU'D LIKE TO PROCESS! *")
    print("*******************************************************************************")
    print("")

    isValidDirectory = False
    containsPlzFiles = False
    while not (isValidDirectory and containsPlzFiles): # loop until both conditions are satisfied
        directory = filedialog.askdirectory(title = "Select the directory containing PLZ files (and PNG thumbnails)") # use tkinter to pop up directory selection dialog

        if directory == "": # user probably hit cancel on file selection dialog
            print("")
            print("...umm, okay. BYE!")
            print("")
            quit() # quit the program completely

        if not ' ' in directory: # check to see if directory has spaces
            isValidDirectory = True # good directory, no spaces
        else: # complain
            isValidDirectory = False # bad directory, contains spaces
            print("")
            print(f"'{directory}' contains spaces.")
            print("Use a directory structure containing no spaces.")
            print("Please try again!")
            print("")
            continue # skip rest of loop and continue to next iteration

        # check to make sure directory contains at least 1 plz file
        if glob.glob(directory + "/*.plz"): # search directory for plz files
            containsPlzFiles = True # we found a plz file, flag this to true
        else: # complain
            containsPlzFiles = False # no plz files found
            print("")
            print(f"'{directory}' does not contain any PLZ files. Choose a directory containing at least 1 PLZ file..")
            print("Please try again!")
            print("")

    print("")
    print(f"You've chosen the directory {directory}")
    print("")

    return directory

def defineDirectories(directory): # function to define directories used globally
    global root_directory, temp_directory, new_file_directory, backup_directory # declare global variables
    root_directory = directory
    temp_directory = os.path.abspath(root_directory + "/FixImages_temp") # define temporary directory name
    new_file_directory = os.path.abspath(temp_directory + "/new_files") # define new file temporary directory name
    backup_directory = os.path.abspath(temp_directory + "/plz_backup") # define backup temporary directory name
    if not os.path.exists(new_file_directory): # make sure this directory doesn't exist (it shouldn't, but let's check anyway)
        print("")
        print(f'creating directory {new_file_directory} for temporary storage of new files.')
        print("")
        os.makedirs(new_file_directory) # recursively create new directory structure <root>/FixImages_temp/new_files
        os.makedirs(backup_directory) # create PLZ backup directory

def checkForThumbnails():
    # check root directory contents, warn user if no thumbnails are found (they can still be moved to the input directory by the user at this point if necessary)
    thumbnailsMissing = False # initialize this flag to false (we'll set to true if we find at least one PLZ archive without a corresponding thumbnail image)
    print(f"Relevant files found in {root_directory}:")
    for path, dirs, files in os.walk(os.path.abspath(root_directory)): # walk through directory and file structure in predefined path to find files
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
        print(f"You can still copy the thumbnail(s) into {root_directory}")
        print("RIGHT NOW if you'd like to include them. Otherwise, they'll need to be added manually afterwards within the tenant.")
        print("*********************************************************************************************************")
        print("")
        input("PLEASE PRESS <ENTER> TO CONTINUE!")
        print("")

def getUserPreferences():
    global publishToDocumoto, cleanup, addHotPointLinks # declare these as global so they can be used elsewhere
    publishToDocumoto = input("Publish the new PLZ pages to Documoto? Type YES to publish: ").upper() == "YES"
    print("")
    addHotPointLinks = input("Would you like to add Hotpoint Links to any of the pages? If so, type YES: ").upper() == "YES"
    print("")
    cleanup = input("Cleanup temporary files after repackaging archives? Type YES to delete: ").upper() == "YES"
    print("")

def backupPLZFiles(): # back up original PLZ archives in temp folder prior to modification
    print("")
    for path, dirs, files in os.walk(os.path.abspath(root_directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            plzBackupFilePath = os.path.join(backup_directory, plzFileName) # join backup path and filename to get absolute file path (for backup)

            # copy the original to the backup directory
            print(f"Backing up {plzFileName}")
            shutil.copyfile(plzFilePath, plzBackupFilePath)

        break # prevent descending into subfolders
    print("")
    print(f"Original PLZ files have been backed up to {backup_directory}")
    print("")
    
def extractArchives():
    # Extract all the contents of zip file in temporary subdirectory
    for path, dirs, files in os.walk(os.path.abspath(root_directory)): # walk through directory and file structure in predefined path to find files
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

def modifyXMLfiles():
    # MODIFY XML and save to new file directory (update name and description attributes in page translation)
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for xmlFileName in fnmatch.filter(files, "*.xml"): # iterate through only the file that match specified extension
            xmlFilePath = os.path.join(path, xmlFileName) # join path and filename to get absolute file path

            pageTitleWithSpaces = getPageTitleStringFromFilename(xmlFileName)
            if pageTitleWithSpaces == "": break # break the loop if zero length list
            
            # use xml.etree to update page attributes 'name' and 'description'
            XMLnamespaces = register_all_namespaces(xmlFilePath, xml_ET) # register namespaces, define a dictionary containing them
            XMLtree = xml_ET.parse(xmlFilePath) # parse xml file into xml tree
            XMLroot = XMLtree.getroot() # get the root tree (in this case, root is <Page>)

            # for <Page>'s <Translation> subelement
            pageTranslationElement = XMLroot.find("Translation", XMLnamespaces) # define Translation element so we can set its attributes
            pageTranslationElement.set('name', pageTitleWithSpaces) # set name
            pageTranslationElement.set('description', pageTitleWithSpaces) # set description

            # for <Page>'s <ProcessingInstructions> subelement
            pageProcessingInstructionsElement = XMLroot.find("ProcessingInstructions", XMLnamespaces) # define <ProcessingInstructions> element so we can set its attributes
            pageProcessingInstructionsElement.set('lockPartTranslations', "true") # lock part translations (to prevent overwriting old part descriptions with new ones.  to overwrite, change this value to "false")

            # if thumbnailFilePath exists, we can safely build and add the <Attachment> element and <Comments> subelement
            # to the root tree. <Comments> contains the filename.  We will need to upload this file also
            # build and add <Attachment> element here:
            thumbnailFileName = xmlFileName.replace('.xml', '.png') # get thumbnail filename using xml filename by replacing xml with png
            thumbnailFilePath = os.path.join(os.path.abspath(root_directory), thumbnailFileName) # join root directory path with thumbnail filename to get thumbnail absolute path
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

            # HOTPOINT LINKING
            if addHotPointLinks: # only perform if user preference is turned on
                print("")
                # here, we need to iterate through each of the <Part> elements to collect data
                # to be used for the new <HotpointLink> elements (which will be children of the
                # root <Page> element).  The following attributes of <HotpointLink> need definitions:
                #   - item (identical to <Part> item attribute)
                #   - locale="en_US"
                #   - targetPageFile (filedialog.askopenfilename() to get PLZ filename, verify it's a PLZ file, THEN replace 'plz' with 'svg')
                #   - targetPageName (use getPageTitleStringFromFilename(targetPageFile) to generate this)
                #   - type="Page"
                for partElement in XMLroot.findall('Part', XMLnamespaces):
                    # define <Part> element's <Translation> subelement to get info from:
                    partTranslationElement = partElement.find("Translation", XMLnamespaces)

                    # get attribute values from each part:
                    partNumber = partElement.get('partNumber') # get part number
                    partDescription = partTranslationElement.get('name') # get part description
                    item = partElement.get('item') # define item number value

                    # let user select PLZ file, strip the path to get only the filename
                    print("")
                    print(f"Please select the PLZ page file you'd like to link to:")
                    print(f"Hotpoint #{item} - {partNumber} - {partDescription}")
                    print("")
                    targetPageFile = ""
                    targetPageFile = os.path.basename(filedialog.askopenfilename(filetypes = [('PLZ files', '*.plz')], title = f"Link to Hotpoint #{item} - {partNumber} - {partDescription}"))

                    if targetPageFile == "": # user hit cancel button on the file selection dialog
                        print("You've cancelled the file selection.")
                        print(f"{partNumber} - {partDescription} WILL NOT BE LINKED!")
                        print("Continuing on to the next part...")
                        print("")
                        continue # continue to the next part element

                    targetPageName = getPageTitleStringFromFilename(targetPageFile)

                    print(f"{partNumber} - {partDescription} will be linked to:")
                    print(targetPageFile)
                    targetPageFile = targetPageFile.replace('.plz', '.svg') # replace plz with svg, as this is what documoto will navigate to
                    print(f"({targetPageFile} in Documoto)")
                    print("")

                    # add <HotpointLink> subelement to <Page> root
                    hotpointLinkElement = xml_ET.SubElement(XMLroot, "HotpointLink")

                    # set attribute values for newly created <HotpointLink> subelement
                    hotpointLinkElement.set('item', item) # the item (hotpoint) number
                    hotpointLinkElement.set('locale', "en_US") # constant value
                    hotpointLinkElement.set('targetPageFile', targetPageFile) # the SVG page file name
                    hotpointLinkElement.set('targetPageName', targetPageName) # the page description
                    hotpointLinkElement.set('type', "Page") # constant value

            # save the modified xml in new file directory
            newXmlFilePath = os.path.join(new_file_directory, xmlFileName)
            XMLtree.write(newXmlFilePath, encoding='utf-8', xml_declaration=True) # need to set xml_declaration to True to preserve first line of xml <?xml version="1.0" encoding="UTF-8"?>

            print(f'saved new xml to:\n{newXmlFilePath}\n')
            
        print("XML ELEMENT AND ATTRIBUTE MODIFICATION COMPLETE!")
        print("")
        break # prevent descending into subfolders

def generateTemporarySVGfiles():
    # use etree to modify attributes in the svg files to prepare them for overlaying onto rasters
    # (these modified SVGs will NOT be repacked into original PLZ archives - they're only for temporary use)
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for svgFileName in fnmatch.filter(files, "*.svg"): # iterate through only the file that match specified extension
            svgFilePath = os.path.join(path, svgFileName) # join path and filename to get absolute file path

            SVGnamespaces = register_all_namespaces(svgFilePath, svg_ET) # register namespaces, define a dictionary containing them
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

def generateNewRasters():
    # use imagemagick mogrify function to overlay new temporary SVG callouts onto original source png raster image
    process = subprocess.Popen(f'mogrify -path {new_file_directory} -format png {temp_directory}/*.svg', # run the shell command, performing mogrify on ALL SVGs in the temp_directory with their corresponding PNG raster images
                           shell=True, stdout=subprocess.PIPE)
    process.wait() # wait for process to finish in current thread before proceeding
    print(f"newly-generated PNGs are located in {new_file_directory}!")
    print("")

def removeSourcePNGandXMLfiles():
    # we need to remove the original png and xml files from the plz archives to avoid duplication
    for path, dirs, files in os.walk(os.path.abspath(root_directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file (if excluded, we don't want to touch it)

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

def repackAndPublishPLZs():
    # now we need to re-pack the new png and xml files into their original archives
    for path, dirs, files in os.walk(os.path.abspath(root_directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file (if excluded, we don't want to touch it)

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

                    thumbnailResponse = requests.request('POST', DOCUMOTO_API_UPLOAD_ENDPOINT_URL, headers = text_headers, files = filesToUpload)
                    printResponseDetails(thumbnailResponse) # print response details

                # now, upload the PLZ
                print("")
                print(f"UPLOADING AND PUBLISHING PLZ: {plzFileName}...")

                filesToUpload = {'file': (plzFileName, open(plzFilePath, 'rb'), 'application/octet-stream'), # use filename as first parameter of 3-tuple
                                 'submitForPublishing': True} # set submitForPublishing to true so file is published once uploaded

                plzResponse = requests.request('POST', DOCUMOTO_API_UPLOAD_ENDPOINT_URL, headers = text_headers, files = filesToUpload)
                printResponseDetails(plzResponse) # print response details

        print("")
        print("PNG AND XML REPACKING INTO ORIGINAL PLZ ARCHIVES IS COMPLETE!")
        if publishToDocumoto: print("ARCHIVES HAVE BEEN UPLOADED AND PUBLISHED TO THE DOCUMOTO TENANT!")
        print("")
        break # prevent descending into subfolders

def cleanupFiles():
    # clean up (delete temporary directory)
    if cleanup:
        shutil.rmtree(temp_directory)
        print(f"deleting temp directory {temp_directory}")
        print("")
    else:
        print(f"keeping temporary directory! files are located in {temp_directory}")
        print("")
    
def fixImagesOuttro():
    print("")
    print("PROCESS IS COMPLETE!!!")
    print("")

    # report exclusion list
    if len(exclude_list) != 0:
        print("***** WARNING *****")
        print("the following PLZ files have been skipped:")
        print("")
        for item in exclude_list:
            print(item)
        print("")
        print("PLEASE PROCESS THESE IN DOCUSTUDIO FIRST, ADDING CALLOUT BUBBLES, AND TRY AGAIN!")
        print("")

    print("============================================")
    print("                 GOODBYE!")
    print("============================================")

def fixImages(): # run functions associated with fixImages process
    imageFixerIntro() # say hi!
    defineDirectories(getRootDirectoryFromUser()) # define global directories based on user-input directory path
    checkForThumbnails() # look for matching thumbnails
    getUserPreferences() # get user preferences for some optional functionality
    backupPLZFiles() # backup the original PLZs prior to modification
    extractArchives() # unpack the PLZs
    modifyXMLfiles() # modify XML elements and attributes
    generateTemporarySVGfiles() # create temp SVG images in preparation for overlay onto rasters
    generateNewRasters() # create the new raster images
    removeSourcePNGandXMLfiles() # wipe out the source PNG and XML files prior to repack
    repackAndPublishPLZs() # repack the new PNG and XML files into their corresponding source archives
    cleanupFiles() # clean up temporary stuff if user wants to
    fixImagesOuttro() # say bye!

def mainMenu(): # main menu structure
    print("MAIN MENU:")

    print("")
    print("1. Fix Images in PLZ file(s)")
    print("0. Exit")
    choice = input("What would you like to do? Enter a number: ")
    if choice == "1": # fix images
        fixImages()
    elif choice == "0": # exit
        exit()
    else:
        print("That's not a valid choice! Try again.")

# END OF FUNCTION DEFINITIONS

# MAIN PROGRAM STARTS HERE

while True: # infinite looping menu
    mainMenu()