#!/usr/local/bin/python
import os, fnmatch, subprocess, shutil, zipfile
from zipfile import ZipFile

def decompressPLZs(directory): # this function decompresses all PLZs in given directory into same directory
    
    temp_directory = os.path.abspath(directory + "/FixImages_temp") # define temporary directory name
    new_png_directory = os.path.abspath(temp_directory + "/new_png") # define new_png temporary directory name

    # Extract all the contents of zip file in temporary subdirectory
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            with ZipFile(plzFilePath, 'r') as archive:
                archive.extractall(temp_directory)
        break # prevent descending into subfolders

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
        break # prevent descending into subfolders

    # now we need to mogrify
    if not os.path.exists(new_png_directory):
        print(f'creating directory {new_png_directory}')
        os.makedirs(new_png_directory) # create new directory (because mogrify is dumb and can't create one itself)

    process = subprocess.Popen(f'mogrify -path {new_png_directory} -format png {temp_directory}/*.svg', # run the shell command
                           shell=True, stdout=subprocess.PIPE)
    process.wait() # wait for process to finish in current thread before proceeding

    # we need to remove the original png source images from the plz archives to avoid duplication
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension
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
        break # prevent descending into subfolders


    # now we need to re-pack the new png files into their original archives
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for plzFileName in fnmatch.filter(files, "*.plz"): # iterate through only the file that match plz file extension

            plzFilePath = os.path.join(path, plzFileName) # join path and filename to get absolute file path
            pngFileName = plzFileName.replace('.plz', '.png')
            pngFilePath = os.path.join(new_png_directory, pngFileName) # remove last 3 characters from plz file path (.plz) to get base filename (no extension).
                                                                                                      # Then, append "png" to it to get the png file path (after joining with new_png_directory path)

            with ZipFile(plzFilePath, 'a') as archive: # append
                archive.write(pngFilePath, arcname=pngFileName) # write the new png file to the archive
                archive.printdir()
        break # prevent descending into subfolders

    # clean up (delete temporary directory)
    cleanup = input("Cleanup temporary files? Type YES to delete.")
    if cleanup.upper() == "YES":
        shutil.rmtree(temp_directory)

print("Welcome to Documoto image fixer!")
print("This is an application to embed Documoto callout bubbles onto the source raster images.")
print("First, extract the svg and png files from plz archives into desired directory.")
print("The directory entered below must include documoto vector images (svg) and")
print("their corresponding raster images (png). The resulting new raster images (png)")
print("will be located in C:\\new_png, and must be re-packaged into the Documoto package (plz)")
print("The original decompressed png and svg files are only used to generate the new raster images")
print("and can be deleted after this process.")

directory = input("Enter directory containing png and svg files:")
directory = directory.replace('\\', '/') # replace backslashes with forward slashes
decompressPLZs(directory)


print("GOODBYE!")