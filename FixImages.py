#!/usr/local/bin/python
import os, fnmatch, subprocess, sys

def findReplace(directory, find, replace):
    filePattern = "*.svg"
    for path, dirs, files in os.walk(os.path.abspath(directory)):
        for filename in fnmatch.filter(files, filePattern):
            filepath = os.path.join(path, filename)
            with open(filepath) as f:
                s = f.read()
            s = s.replace(find, replace)
            with open(filepath, "w") as f:
                f.write(s)

def mogrify(directory):
    process = subprocess.Popen(f'mogrify -path C:/new_png -format png {directory}/*.svg', # shell command is: mogrify -path C:\new_png -format png *.svg
                           shell=True, stdout=subprocess.PIPE)
    process.wait()

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