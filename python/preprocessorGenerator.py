#python3
# This script is used to generate preprocessor for a list of test case. 
# Input: tcDir (folder contains test case for execution)
#		stubDir: Folder contains output file AMSTB_Preproccessor.h
# Output: AMSTB_Preproccessor.h
import os, re, fnmatch
def main():
    # Path to folder contain test casse for execution
    tcDir="F:\\work\\3_5G_Soft_2\\branches\\eTAS\\04_Output\\01_Test_Spec\\02_Test_Case"
    # path to stub for compileing
    stubDir='F:\\work\\eTAS\\Source\\Source_201807xx\\201807xx\\sorce\\eTAS_SourceCode_build\\Sources\\stub'
    file=stubDir + '\\AMSTB_Preprocessor.h'
    data = find_files(tcDir,'*.csv')
    with open(file, 'w') as f:
        f.write(data)
# The function return list of file without extension and path
def find_files(directory, pattern):
    strData =''
    for root, dirs, files in os.walk(directory):
        for basename in files:
            if fnmatch.fnmatch(basename, pattern):
                print (basename)
                filename = basename.split('.')[0]
                filename ='#define ' + filename.upper() +'\n'
                strData += filename
    return strData
if __name__ == '__main__':
    main()
