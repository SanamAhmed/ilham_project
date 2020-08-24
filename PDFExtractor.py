from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
import pandas as pd
#from DBHandler import DBHandler
import re
import docx
import spacy
#import Translator
from nltk.corpus import stopwords
stop_words = set(stopwords.words('english'))
from nltk.tokenize import word_tokenize

print(stop_words)
from pycorenlp import *
import time
import traceback
import json
def getDataPDF(filename):
    try:
        filenametxt=filename.replace("pdf","txt")
        fp = open(filename, 'rb')
        parser = PDFParser(fp)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        laparams.word_margin = 1.0
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        extracted_text = ''
        with open(filenametxt, "a", encoding="utf-8") as f:
            for page in doc.get_pages():
                interpreter.process_page(page)
                layout = device.get_result()
                fullText = ""
                extracted_text = ""
                for lt_obj in layout:
                    if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                        extracted_text += lt_obj.get_text()


                n1 = extracted_text.replace("\t", " ")
                n2 = n1.replace("\r", " ")
                #n3 = n2.replace("\n", " ")
                finaltext = n2.replace("\u00a0", " ")
                f.write(finaltext+"\n")
                fullText = fullText + finaltext


             #   print("Text Extracted")
            #    print(fullText)
    except Exception as e:
       # print("Exception in Processing file" + self.fname + str(e))
        traceback.print_exc()

    return True

#getDataPDF("D:\Crymzee\IlhamProject\IlhamProject1.0\DR-03122019 (1).pdf")

def extractMemberWiseConversation(filename):
    print("In extracting conversation")
    f = open(filename, 'rb')
    data = f.read().decode('utf8', 'ignore')
    lines = data.split("\n")
    memberData=[]
    index =0
    lastindex=-1
    member = None
    for line in lines:
        if(":" in line and index==0 and member ==None):
            #print("found member")
            namelist=line.split(":")
            member=[]
            if len(namelist[1])== 1:
                continue
            member.append(namelist[0])
            member.append(namelist[1])
            #print("memberName")
            lastindex=index
            index=index+1
        elif(":" in line and index > 0):
            #print("found new member")
            if(len(member)>1):
                memberData.append(member)
            #print("reinitializing member")
            namelist = line.split(":")
            member = []
            member.append(namelist[0])
            member.append(namelist[1])
            #print("memberName", member[0])
            lastindex = index
            index = index + 1
        elif (member != None and member!=[]):
            member.append(line)

        #print(line)


    return memberData

def extractMembersSection(filename):
    f = open(filename, 'rb')
    data = f.read().decode('utf8', 'ignore')
    lines = data.split("\n")
    memberData = []

    linenumbers = 0
    listmembersConvos = []
    for line in lines:
        if (len(line) <= 4 and line.__contains__(".")):
            print(line)
            listmembersConvos.append(linenumbers)
        linenumbers = linenumbers + 1

    for linenumber in listmembersConvos:
        member = []
        member=getMemberConvo(lines,linenumber)
        if (member !=[]):
            memberData.append(member)
    return memberData

def getMemberConvo(lines,linenumber):
    i=linenumber+1
    line = lines[i]
    #print("Member Name ",line)
    member=[]
    if line.__contains__("]"):
        memberNameList = line.split("]")
        member.append(memberNameList[0]+"]")
        member.append(memberNameList[1])
    elif line.__contains__("minta"):
        memberNameList = line.split("minta")
        member.append(memberNameList[0])
        member.append(memberNameList[1])
    i=i+1
    while i < len(lines):
        if(member!=[]):
            line=lines[i]
            if(line.__contains__(":")):
                spilitlist = line.split(":")
                if(len(spilitlist[1])>3):
                   break
            else:
                #print("Member:",member[0],"said",line)
                member.append(line)
        i=i+1
    return member

def word_count_perMember(member,stopwordfile):
    counts = dict()
    #counts.update("MemberName",member[0])
    i=1
    columnnames=["stopword"]
    data = pd.read_excel(stopwordfile)
    stopword_df = pd.DataFrame(data, columns=['stopwords'])
    stopwords_indo = stopword_df["stopwords"].values
    #print(stopwords_indo)
    #print(stopword_df)
    while i < len(member):
        line=member[i]
        words = word_tokenize(line)
        for word in words:
            if word not in stopwords_indo :
                if word in counts:
                    counts[word] += 1
                else:
                    counts[word] = 1

        i=i+1
    counts.update(MemberName= member[0])
    return counts

def memberWiseWordCount(memberDataConvo1,filename):
    for member in memberDataConvo1:
        member_word_count = word_count_perMember(member,filename)
        print(member_word_count)


'''
def search(pattern1, pattern2 , filename):
    memberDataConvo1 = extractMemberWiseConversation(filename)
    memberDataConvo2 = extractMembersSection(filename)
    memberDataConvo1.extend(memberDataConvo2)
    pattern1list = pattern1.split(",")
    pattern2list = pattern2.split(",")
    for memberName in pattern1list:
        for keyword in pattern2list:
            for member in memberDataConvo1:
'''

memberDataConvo1 = extractMemberWiseConversation("D:\Crymzee\IlhamProject\IlhamProject1.0\DR-03122019 (1).txt")
memberDataConvo2 = extractMembersSection("D:\Crymzee\IlhamProject\IlhamProject1.0\DR-03122019 (1).txt")
memberDataConvo1.extend(memberDataConvo2)
'''
for member in memberDataConvo1:
    # print(member)
    if (member != None and len(member) > 1):
        print("memberName", member[0])
        i = 1
        while i < len(member):
            print(member[0], "said:", member[i])
            i = i + 1
'''

memberWiseWordCount(memberDataConvo1,"D:\Crymzee\IlhamProject\IlhamProject1.0\stopwordbahasa.xlsx")


