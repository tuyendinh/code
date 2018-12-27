#!python3
from bs4 import BeautifulSoup
from html_table_extractor.extractor import Extractor
import sys, glob, os, re, codecs, fnmatch, pprint
from collections import defaultdict

def main ():
    if  len(sys.argv) > 3:
        print("argument execution")
        rootDir = sys.argv[1]
        manualFile = sys.argv[2]
        stubFile = sys.argv[3]
    elif len(sys.argv) > 2:
        rootDir = sys.argv[1]
        stubFile = sys.argv[2]
    else:
        # Path to folder contain sub functions that are exported by WINAMS. Folder that contain *_vmod.html file (e.g Apl_Config_GetAplChec&0001_vmod.html)
        rootDir="S:\\04_CasePlayer\\eTAS_PBW\\HTML"
        # Path to stub file
        stubFile='S:\\06_SourceCode\\Stub\\AMSTB_SrcFile.c'
        # File that contain missing sub function
        #manualFile='F:\\work\\3_5G_Soft_2\\branches\\eTAS\\04_Output\\05_Source_Code\\04_Stub_Functions\\Caller.txt'
        manualFile=''
    dictCallee=defaultdict(list)
    dictCallee = ParseSubFunctions(dictCallee,rootDir)
    if os.path.isfile(manualFile):
        dictCallee = getCalleeFromFile(dictCallee,manualFile)
    #pprint.pprint(dictCallee)
    updateStub(dictCallee, stubFile)
def getCalleeFromFile(dictCallee, file):
    with open(file, 'r') as fileObj:
        lines=fileObj.readlines()
    for line in lines:
        #if not comment
        if line.find(';')==-1:
            str1=line.split(':')[0]
            str2=line.split(':')[-1]
            arrStr2 = str2.split(',')
            for temp in arrStr2:
                temp = temp.strip()
                temp = 'defined('+ 'TEST_' + temp.upper() + ')'
                if(temp not in dictCallee[str1]):
                    dictCallee[str1].append(temp)
    return dictCallee
def updateStub(dictCallee, stubFile):
    newStubFile=stubFile + '_new'
    PreprocessorFile = os.path.dirname(os.path.realpath(stubFile)) + "\\AMSTB_Preprocessor.h_new"
    Newlines=[]
    # Find name of the function is stub file
    funRegex=re.compile(r'''\/\*\s*[A-Z_0-9]+\[[a-zA-Z_0-9]+\.(c|dbg):([a-zA-Z_0-9]+):.+\*\/''')
    with open(stubFile,'r') as fileObj:
        lines =  fileObj.readlines()
    i = 0
    while(i <  len(lines)):
        m = funRegex.match(lines[i])
        if m:
            # Create preprocessor
            if m.groups()[1] in dictCallee:
                pre='||'.join(dictCallee[m.groups()[1]])
            else:
                pre = '0'
            pre ='#if ' + pre + '\n'
            Newlines.append(pre + lines[i])
            # Start read until complete function
            countBracket = 0
            flag = 1
            while (countBracket != 0 or flag == 1):
                i +=1
                #print (lines[i])
                if '{' in lines[i]:
                    flag = 0
                    countBracket +=1
                if '}' in lines[i]:
                    countBracket -=1
                Newlines.append(lines[i])
            Newlines.append('#endif' + '\n')
        else:
            Newlines.append(lines[i])
        i +=1
    # Create new stub file
    text = ''.join(Newlines)
    with open(newStubFile,'w') as fileObj:
       fileObj.write(text)
    # Create preprocessor file
    regex = re.compile(r'''#if\s+((\|\|)?defined\s?\([a-zA-Z0-9_]+\))+''')
    matches = regex.findall(text, re.MULTILINE)
    text =[]
    for group in matches:
        text.append(''.join(group))
    #remove duplicate
    newText=[]
    for item in text:
        if item not in newText:
            newText.append(item)
    newText='\n'.join(newText)
    newText=newText.replace('||','')
    newText=newText.replace(' ','')
    newText=newText.replace('(',' ')
    newText=newText.replace(')','')
    newText=newText.replace('defined','#define ')
    with open(PreprocessorFile,'w') as fileObj:
       fileObj.write(newText)


def ParseSubFunctions(dictCallee,root):
    htmlFiles = find_files(root,"*_vmod.html")
    #file ="F:\\work\\3_5G_Soft_2\\branches\\eTAS\\02_Input_Data_From_Onsite_Team\\Subfunctions\\Application\\APL_SysM\\APL_SysM_CreateTrigger.html"
    # hashtable to store funcion call
    for file in htmlFiles:
        #print(file)
        with codecs.open(file,'r',encoding='Shift_JIS') as f:
        #with codecs.open(file,'r',encoding='ISO-8859-1') as f:
            text = f.read()
        soup=BeautifulSoup(text,'html.parser')
        # Get function name
        FuncName=soup.findAll("table")[1]
        FuncName = Extractor(str(FuncName)).parse().return_list()[0][1]
        FuncName = ''.join(['defined(','TEST_', FuncName.upper(),')'])
        #print(FuncName)
        # Get sub function list
        subFuncTable=soup.findAll("table")[-2]
        extractor = Extractor(str(subFuncTable))
        extractor.parse()
        functionList = extractor.return_list()
        if (('関数名' in functionList[1]) or ('Function' in functionList[1])):
            for i in range(2,len(functionList)):
                #print (functionList[i][0])
                dictCallee[functionList[i][0]].append(FuncName)
    return dictCallee

def find_files(directory, pattern):
    fileList=[]
    for root, dirs, files in os.walk(directory):
        for basename in files:
            if fnmatch.fnmatch(basename, pattern):
                filename = os.path.join(root, basename)
                print (filename)
                fileList.append(filename)
    return fileList
if __name__ == '__main__':
    main()

