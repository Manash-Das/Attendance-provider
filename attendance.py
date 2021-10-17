import glob
import os
import cv2
import pandas as pd
import os.path, time
from difflib import SequenceMatcher
import pytesseract
import matplotlib

mpl_data_dir = matplotlib.get_data_path()
datas = [ 
    (mpl_data_dir, "matplotlib/mpl-data"), 
]
# pytesseract.pytesseract.tesseract_cmd=r'C:\Users\MANASH DAS\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def imageCollection():
    x=os.getcwd()                       ## get directory of file
    typesOfPhoto=[]                     ## list of photo extension
    AllFiles=[]                         ## list of all photo with that extension
    listOfPhotos=[]                     ## list of all photo name
    typesOfPhoto.append(x+"\\*.jpg")    ## added jpg extension
    typesOfPhoto.append(x+"\\*.jpeg")   ## added jpeg extension
    typesOfPhoto.append(x+"\\*.png")    ## added png extension
    for i in typesOfPhoto:
        AllFiles.append(glob.glob(i))   ## getting all photos with that extension
    for i in AllFiles:
        for j in range(len(i)):
            listOfPhotos.append(AllFiles[AllFiles.index(i)][j].split("\\")[-1]) ## get the photo name
    return listOfPhotos                 ## return all photos name

def imageProcessing(imageLocation):
    #Loadinng image
    attendanceImage=cv2.imread(imageLocation)
    #Resize image
    resizeImage=cv2.resize(attendanceImage,(0,0),None,2,2,cv2.INTER_LINEAR_EXACT)
    #Converting image to gray image
    grayImage=cv2.cvtColor(resizeImage,cv2.COLOR_BGR2GRAY)
    #Creating blur in the image
    blurImage=cv2.bilateralFilter(grayImage,5,40,40)
    #Threshold image
    thresholdImage = cv2.adaptiveThreshold(blurImage,155,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,5,5)
    return thresholdImage

def imageToTextConverter(attendanceThreshold):
    ##converting image to text
    return(pytesseract.image_to_string(attendanceThreshold).upper())
    
def removeUnwanted(text):
    ## unwanted text
    unrequiredChar="1234567890~!`@#$%^&*()-_=+}]|\{[:;"'<,>.?/'"]n«©£€}" 
    for char in unrequiredChar:
        ##removed character
        text=text.replace(char," ")
    ## cleared text list                                     
    return text.split()

def filteringText(listOfWords,firstName):
    wordsIndex=0
    namesFound=[]
    while(wordsIndex<len(listOfWords)):
        name=""
        if listOfWords[wordsIndex] in firstName:
            firstNameIndex=firstName.index(listOfWords[wordsIndex])
            if listOfWords[wordsIndex+1]==firstName[firstNameIndex+1]:
                name=listOfWords[wordsIndex]+" "+listOfWords[wordsIndex+1]
                del listOfWords[wordsIndex]
                del listOfWords[wordsIndex]
                del firstName[firstNameIndex]
                del firstName[firstNameIndex]
                if listOfWords[wordsIndex]==firstName[firstNameIndex]:
                    name=name+ " " +listOfWords[wordsIndex]
                    del listOfWords[wordsIndex]
                    del firstName[firstNameIndex]
                namesFound.append(name)
            else:
                wordsIndex+=1
        else:
            wordsIndex+=1
    return namesFound

def secondFilteringText(listOfWords,firstName):
    namesFound=[]
    for lenWord,word in enumerate(listOfWords):
        for lenFirstName,name in enumerate(firstName):
            if SequenceMatcher(None,word,name).ratio()>0.7:
                if SequenceMatcher(None,listOfWords[lenWord+1],firstName[lenFirstName+1]).ratio()>0.7:
                    namesFound.append(f"{name} {firstName[lenFirstName+1]}")
    return namesFound

def Name():
    firstName=["MANASH","DAS","","TANUMAY","DAS","","ROHIT","KUMAR", "SAH","","MAMON","TALAFDAR","","KAUSTAV","MANDAL",
               "","SOURAV","PAUL","","HIMANSHU","SINGH","","SANDIP","MANDAL","","ARIJIT","ADHIKARY","","SANCHITA","DAS",
               "","RAMCHANDRA","GHOSH","","RAJESWAR","SAMANTA","","ADITYA","SHEE","","SOUVIK","KUMAR","MONDAL","",
               "PRAMIT","DATTA","","JABED","RAHAMAN","","ADARSH","KUMAR","","PRAGATI","TANDON","","ABHISHEK","PAUL","",
               "MEGHA","BISWAS","","ARGHYADEEP","DEBNATH","","PAYEL","MONDAL","","DIPANKAR","TARAFDAR"]
    return firstName

studentNames=["MANASH DAS","TANUMAY DAS","ROHIT KUMAR SAH","MAMON TALAFDAR","KAUSTAV MANDAL","SOURAV PAUL","HIMANSHU SINGH",
       "SANDIP MANDAL","ARIJIT ADHIKARY","SANCHITA DAS","RAMCHANDRA GHOSH","RAJESWAR SAMANTA","ADITYA SHEE","SOUVIK KUMAR MONDAL",
       "PRAMIT DATTA","JABED RAHAMAN","ADARSH KUMAR","PRAGATI TANDON","ABHISHEK PAUL","MEGHA BISWAS","ARGHYADEEP DEBNATH",
       "PAYEL MONDAL","DIPANKAR TARAFDAR"]

def updateInExcel(studentNames,NamesFoundList,imageName):
    try: 
        studentData=pd.read_excel("attendanceList.xlsx")
    except: 
        studentData=pd.DataFrame()
        studentData["SL. No"]=list(range(1,len(studentNames)+1))
        studentData["NAMES"]=studentNames
        
    markers=[0 for i in range(len(studentNames))]
    for i in range(len(studentNames)):
        if studentNames[i] in NamesFoundList:
            markers[i]=1
    
    times=time.ctime(os.path.getctime(imageName)).split()
    Date=times[0] +" " + times[2]+ " "+times[1]+" "+times[4]
    print(Date)
    if Date in list(studentData.columns):
        dateAttendance=studentData[Date]
        for index in range(len(dateAttendance)):
            markers[index]=markers[index]|dateAttendance[index]
        
    
    studentData[Date]=markers
    studentData.to_excel("attendanceList.xlsx",index=False)





listOfPhotos=imageCollection()
for photo in listOfPhotos:
    image=imageProcessing(photo)
    text=imageToTextConverter(image)
    clearedText=removeUnwanted(text)
    firstName=Name()
    namesFound=filteringText(clearedText, firstName)
    namesFound.extend(secondFilteringText(clearedText,firstName))
    print(namesFound)
    updateInExcel(studentNames, namesFound, photo)
# print(list(text))
    