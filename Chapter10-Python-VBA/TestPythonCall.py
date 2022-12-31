import sys
import os

#Arguments passed from Macro
vTextFilePath= sys.argv[1]
vTextToWrite=sys.argv[2]

vTextFilePath=vTextFilePath.replace(os.sep,"/")

#Write a Text File
with open(vTextFilePath, 'w') as f:
    f.write(vTextToWrite)