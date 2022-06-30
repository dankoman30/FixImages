#!/usr/local/bin/python
import os, fnmatch, subprocess, shutil, zipfile
from zipfile import ZipFile

exclude_list = []

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
    new_png_directory = os.path.abspath(temp_directory + "/new_png") # define new_png temporary directory name

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

    # now that all files are extracted, let's find and replace
    for path, dirs, files in os.walk(os.path.abspath(temp_directory)): # walk through temp_directory to find files
        for svgFileName in fnmatch.filter(files, "*.svg"): # iterate through only the file that match specified extension
            svgFilePath = os.path.join(path, svgFileName) # join path and filename to get absolute file path
            with open(svgFilePath) as f:
                s = f.read() # open the file
            s = s.replace('stroke=\"#FFFFFF\" stroke-width=\"2\" ', '') # find and replace
            s = s.replace('<text', '<text font-size=\"36px\" dy=\"9px\"') # find and replace
            with open(svgFilePath, "w") as f:
                f.write(s) # close the file
            print(f"performing find-and-replace on {svgFileName}...")
        break # prevent descending into subfolders
        print("find-and-replace complete!")
        print("")

    # now we need to mogrify
    if not os.path.exists(new_png_directory):
        print("")
        print(f'creating directory {new_png_directory} for temporary storage of new PNG files.')
        print("")
        os.makedirs(new_png_directory) # create new directory (because mogrify is dumb and can't create one itself)

    process = subprocess.Popen(f'mogrify -path {new_png_directory} -format png {temp_directory}/*.svg', # run the shell command
                           shell=True, stdout=subprocess.PIPE)
    process.wait() # wait for process to finish in current thread before proceeding
    print(f"newly-generated PNGs are located in {new_png_directory}!")
    print("")

    # we need to remove the original png source images from the plz archives to avoid duplication
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
                if (item.filename[-4:] != '.png'): # write all except png files to new archive
                    zout.writestr(item, buffer)
            zout.close() # close files
            zin.close() # close files
            os.remove(old_filepath) # delete old file
            os.rename(new_filepath, old_filepath) # rename new file to old filename
            print(f"removing source PNG from original archive {plzFileName}")
        break # prevent descending into subfolders
        print("source PNG removal from original PLZ archives is complete!")
        print("")


    # now we need to re-pack the new png files into their original archives
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            if isExcluded(plzFileName): # check to see if plz file is in excluded list, before doing anything else
                continue # continue to next loop iteration, skipping this file

            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            pngFileName = plzFileName.replace('.plz', '.png')
            pngFilePath = os.path.join(new_png_directory, pngFileName) # remove last 3 characters from plz file path (.plz) to get base filename (no extension).
                                                                       # Then, append "png" to it to get the png file path (after joining with new_png_directory path)
            print("")
            print(f"re-packing new PNG into archive {plzFileName}")
            print("ARCHIVE CONTENTS:")
            with ZipFile(plzFilePath, 'a') as archive: # append
                archive.write(pngFilePath, arcname=pngFileName) # write the new png file to the archive
                archive.printdir()

        break # prevent descending into subfolders
        print("")
        print("PNG re-packing is completed!")
        print("")

    # clean up (delete temporary directory)
    print("")
    print("PROCESS IS COMPLETE!!!")
    print("")
    cleanup = input("Cleanup temporary files? Type YES to delete: ")
    print("")
    if cleanup.upper() == "YES":
        shutil.rmtree(temp_directory)
        print(f"deleting temp directory {temp_directory}")
        print("")
    else:
        print(f"keeping temporary directory! files are located in {temp_directory}")
        print("")
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

fixTheFiles(directory) # FIX THE FILES!


print("GOODBYE!")