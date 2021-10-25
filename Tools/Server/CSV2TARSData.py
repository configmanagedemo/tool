# -*- coding:gbk -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import JceStructPrinter
import logging
import os.path
import importlib
import os
import csv
import codecs
import shutil
import platform

from tars import core as tarscore

class CSVFile(object):
    def __init__(self, dataPath, stName, structName = ''):
        self.filepath = os.path.join(dataPath, stName + ".csv")
        if len(structName) is 0:
            self.stName = stName
        else:
            self.stName = structName



    def loadCSVData(self, mapData):
        err_str=""
        if not os.path.exists(self.filepath):
            err_str="not exist file:" + self.filepath
            return err_str
        csvData = []
        csvfile = open(self.filepath, 'r')
        csv_reader = csv.reader(csvfile)
        for row in csv_reader:
            csvData.append(row)
        csvType = csvData[0]
        csvField = csvData[1]

        count = 0
        for item in csvData[2:]:
            count += 1
            singleItem = self.cls()

            for index in range(len(csvType)):
                try:                    
                    if csvType[index].strip() == 'INT':
                        if item[index].strip() == '':
                            setattr(singleItem, csvField[index], 0)
                        else:
                            setattr(singleItem, csvField[index].strip(), int(float(item[index].strip())))
                    elif csvType[index].strip() == 'FLOAT':
                        if item[index].strip() == '':
                            setattr(singleItem, csvField[index], 0)
                        else:
                            setattr(singleItem, csvField[index].strip(), float(item[index].strip()))
                    elif csvType[index].strip() == 'STRING':
                        setattr(singleItem, csvField[index].strip(), item[index].strip().decode('gbk').encode('utf-8'))
                    elif csvType[index].strip() == 'BOOL':
                        if item[index].strip() == '':
                            setattr(singleItem, csvField[index], 0)
                        else:
                            setattr(singleItem, csvField[index].strip(), int(float(item[index].strip())))
                    else:
                        err_str = "unknown type:" + csvType[index] + ", File:" +self.filepath
                        
                except Exception as err:
                    err_str = "file:%s Name:%s index:%d line:%d fieldname:%s" % (self.filepath, self.stName, index, count, csvField[index]) +"\n err:" + err
                    return err_str
                    
            key = getattr(singleItem, csvField[0].strip())
            if mapData.has_key(key):
                err_str ="Duplicate key found. File:" + self.filepath + ", key:" + str(key)
                return err_str
            else:
                mapData[key] = singleItem
        return err_str

class DotBFile(object):

    def __init__(self, filename, stName, clsName):
        self.filename = filename
        self.csvFiles = []
        self.stName = stName
        #self.validator = validator
        self.allDataTmp = None
        self.mod = importlib.import_module(stName)
        self.cls = getattr(self.mod, clsName)

    def generateData(self):
	    #allData = getattr(self.mod, self.stName)()
        allData = self.cls()
        for csvFile in self.csvFiles:
            dataMap = getattr(allData, 'map' + csvFile.stName)
            csvFile.cls = getattr(self.mod, 'T' + csvFile.stName)
            err_str = csvFile.loadCSVData(dataMap)
            if(err_str!=""):
                print('\033[1;41m' + err_str +'\033[0m')
                return 1

        outfile = open(self.filename, 'wb')

        packageStream = tarscore.TarsOutputStream()
        allData.writeTo(packageStream, allData)
        outfile.write(packageStream.getBuffer())
        outfile.close()
        self.printData()
        return 0

    def printData(self):
        outfile = open(self.filename, 'rb')
        outfilelog = open(self.filename + ".log", 'w')

        allData = self.cls()
        packageStream = tarscore.TarsInputStream(outfile.read())
        allData = allData.readFrom(packageStream)

        logstr = JceStructPrinter.printStructToString(allData)
        outfilelog.write(logstr)
        outfilelog.close()
        
    def loadData2Local(self):

        self.allData = globals()[self.stName]()
        for csvFile in self.csvFiles:
            dataMap = getattr(self.allData, 'map' + csvFile.stName)
            csvFile.loadCSVData(dataMap)

      

    def writeLocal2File(self):
        outfile = open(self.filename, 'wb')

        packageStream = tarscore.TarsOutputStream()
        self.allData.writeTo(packageStream, self.allData)
        outfile.write(packageStream.getBuffer())
        outfile.close()
        self.printData()        
        

