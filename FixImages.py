#!/usr/local/bin/python
import os, fnmatch, subprocess, sys

# =====================================================================================================================
# FUNCTIONS:
# =====================================================================================================================
def findReplace(directory, find, replace): # this function takes directory path, string to find, and string to replace with, as parameters
    filePattern = "*.svg"
    for path, dirs, files in os.walk(os.path.abspath(directory)): # walk through directory and file structure in predefined path to find files
        for filename in fnmatch.filter(files, filePattern): # iterate through only the file that match specified extension
            filepath = os.path.join(path, filename) # join path and filename to get absolute file path
            with open(filepath) as f:
                s = f.read() # open the file
            s = s.replace(find, replace) # replace instances of 'find' string with 'replace' string
            with open(filepath, "w") as f:
                f.write(s) # close the file

def mogrify(directory): # this function runs mogrify to write the svg (with embedded png) as new raster image
                        # IMPORTANT: IT'S ASSUMED THAT MOGRIFY (PART OF IMAGEMAGICK) IS ALREADY INSTALLED!!!
    process = subprocess.Popen(f'mogrify -path C:/new_png -format png {directory}/*.svg', # run the shell command
                           shell=True, stdout=subprocess.PIPE)
    process.wait() # wait for process to finish in current thread before proceeding

# =====================================================================================================================
# =====================================================================================================================





# =====================================================================================================================
# MAIN PROGRAM BODY
# =====================================================================================================================
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
findReplace(directory, 'stroke=\"#FFFFFF\" stroke-width=\"2\" ', '') # remove while outlines from bubbles
findReplace(directory, '<text', '<text font-size=\"36px\" dy=\"9px\"') # increase text size and center text in bubble
mogrify(directory)

print("GOODBYE!")