import xlrd
import time
import os
import sys
reload(sys)
sys.setdefaultencoding('utf8')   

ExcelPath="../../Excels/"
CharacterList = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_1234567890"
CharacEnumNum = "-1234567890."
CharacEnumBoo = "10."

ColumTypeList = ['INT','STRING','BOOL','FLOAT','TEXT']
ValidSignList = ['*','*#','#','#*','']

ExcelNameList = []
ExcelDataList = []

#定位字串
def StrESRC(ExcelName,SheetName,row,col):
    StrCol = ""
    while True:
        a = col/26
        if a > 0:
            StrCol += CharacterList[a-1]
            col %= 26
        else:
            StrCol += CharacterList[col-1]
            break
    return ExcelName + "|"+ SheetName + " row:" + str(row) + " col:" + StrCol + " "

#样式日志
def LogError(LogStr):
    os.system('')
    try : print ("\033[1;37;41m"+"ERROR:"+"\033[0m"+" \033[1;31;40m"+LogStr+"\033[0m")
    except :
        LogError("LogError Catch Error")

def LogWarning(LogStr):
    os.system('')
    try : print ("\033[1;31;40m"+LogStr+"\033[0m")
    except :
        LogError("LogError Catch Error")
    
def LogInfo(LogStr):
    os.system('')
    try : print ("\033[1;32;40m"+LogStr+"\033[0m")
    except :
        LogError("LogInfo Catch Error")

def LogFinal(LogStr):
    os.system('')
    try : print ("\033[1;34;40m"+LogStr+"\033[0m")
    except :
        LogError("LogFinal Catch Error")

#取文件列表
def GetExcelFiles():
    if not os.path.exists(ExcelPath):
        LogError(ExcelPath + " not exist")
        return False
    for ita in os.listdir(ExcelPath):
        if ita[0] == "~":
            continue
        if ".svn" in ita:
            continue
        ExcelNameList.append(ita)
    if len(ExcelNameList) == 0:
        LogError("has no files in " + ExcelPath)
        return False
    return True

#文件类型检查
def CheckFileType():
    for ita in ExcelNameList:
        if ".xlsm" not in ita[-5:]:
            LogError(ita + " is not .xlsm")
            return False
    return True

#文件名检查
def CheckFileName():
    for ita in ExcelNameList:
        for c in ita[:-5]:
            if c not in CharacterList:
                LogError("File name " + ita + "|" + ita[:-5] + " not in a-z、A-Z、_、1-0")
                return False
    return True

#Excel读取测试
def CheckLoadXlsm():
    DetailStr = ""
    for xlsm in ExcelNameList:
        DetailStr += ExcelPath+xlsm +"|"
        try : data = xlrd.open_workbook(ExcelPath+xlsm)
        except:
            LogError(xlsm+" xlrd.open_workbook catch error!")
            return False
        ExcelDataList.append(data)
        try : names = data.sheet_names()
        except:
            LogError(xlsm+" data.sheet_names() catch error!")
            return False
        for name in names:
            try : sheet = data.sheet_by_name(name)
            except:
                LogError(xlsm+"|" + name + " data.sheet_by_name(name) catch error!")
                return False
            DetailStr += name + ","+str(sheet.nrows)+","+str(sheet.ncols) + "|"
        DetailStr += "\n"
    #LogInfo(DetailStr)
    return True

#INDEX存在检查
def CheckShtIndex():
    for i in range(len(ExcelDataList)):
        names = ExcelDataList[i].sheet_names()
        if names[0] != "INDEX":
            LogError(ExcelNameList[i]+": can't find name 'INDEX' of sheet")
            return False
    return True

#INDEX格式检查
def CheckIndexFmt():
    for i in range(len(ExcelDataList)):
        sheet = ExcelDataList[i].sheet_by_name("INDEX")
        if sheet.ncols <= 5:
            LogError(ExcelNameList[i] + "|INDEX: content incomplete!")
            return False
        for irow in range(2,sheet.nrows):
            if sheet.cell(irow,0).value != "":
                if sheet.cell(irow,2).value == "":
                    LogError(ExcelNameList[i] + "|INDEX: row:"+str(irow+1) +" col:D is empty")
                    return False
                if sheet.cell(irow,5).value == "":
                    LogError(ExcelNameList[i] + "|INDEX: row:"+str(irow+1) +" col:G is empty")
                    return False
    return True

#INDEX内容检查
def CheckIndexCnt():
    for i in range(len(ExcelDataList)):
        sheet = ExcelDataList[i].sheet_by_name("INDEX")
        for irow in range(2,sheet.nrows):
            sheet_name = sheet.cell(irow,2).value
            if sheet_name != "":
                for c in sheet_name:
                    if c not in CharacterList:
                        LogError(ExcelNameList[i]+"|"+sheet_name + "|"+ sheet_name + " not in a-z、A-Z、_、1-0")
                        return False
                if sheet_name not in ExcelDataList[i].sheet_names():
                    LogError(ExcelNameList[i]+"|"+sheet_name + ": config sheet name not exist!")
                    return False
    return True

#Sheet名重名检查
def CheckSheetNms():
    TotalSheetNames = []
    TotalExcelNames = []
    for i in range(len(ExcelDataList)):
        INDEX = ExcelDataList[i].sheet_by_name("INDEX")
        for irow in range(2,INDEX.nrows):
            if INDEX.cell(irow,0).value == "":
                continue
            SheetName = str(INDEX.cell(irow,2).value)
            for pos in range(len(TotalSheetNames)):
                if SheetName == TotalSheetNames[pos]:
                    LogError(TotalExcelNames[pos] + "|"+TotalSheetNames[pos] + " -- " + ExcelNameList[i] + "|"+SheetName + " same sheet name not allowed!")
                    return False
            TotalSheetNames.append(SheetName)
            TotalExcelNames.append(ExcelNameList[i])
    #print(TotalSheetNames)
    #print(TotalExcelNames)
    return True

#Sheet格式检查
def CheckSheetFmt():
    for i in range(len(ExcelDataList)):
        INDEX = ExcelDataList[i].sheet_by_name("INDEX")
        for irow in range(2,INDEX.nrows):
            if "*" not in INDEX.cell(irow,0).value:
                continue
            SheetName = INDEX.cell(irow,2).value
            ExcelName = ExcelNameList[i]
            sheet = ExcelDataList[i].sheet_by_name(SheetName)
            #print(ExcelName,SheetName)

            #标记行检查
            row_0 = sheet.row(0)
            for isign in range(1,len(row_0)):
                if str(row_0[isign].value) not in ValidSignList:
                    LogError(StrESRC(ExcelName,SheetName,1,isign+1)+str(row_0[isign].value) + " not in "+str(ValidSignList))
                    return False

            #标记列检查
            col_0 = sheet.col(0)
            for isign in range(1,len(col_0)):
                if str(col_0[isign].value) not in ValidSignList:
                    LogError(StrESRC(ExcelName,SheetName,isign+1,1)+str(col_0[isign].value) + " not in "+str(ValidSignList))
                    return False

            #列类型检查
            row_1 = sheet.row(1)
            for isign in range(1,len(row_1)):
                if str(row_1[isign].value) not in ColumTypeList and "*" in str(sheet.cell(0,isign).value):
                    LogError(StrESRC(ExcelName,SheetName,2,isign+1)+str(row_1[isign].value)+" not in "+str(ColumTypeList))
                    return False
            
            #类型字符合法性检查
            for icol in range(1,sheet.ncols):
                if "*" not in str(sheet.cell(0,icol).value):
                        continue
                CharList = []
                if str(sheet.cell(1,icol).value) == "INT":
                    CharList = CharacEnumNum
                if str(sheet.cell(1,icol).value) == "BOOL":
                    CharList = CharacEnumBoo
                if str(sheet.cell(1,icol).value) == "FLOAT":
                    CharList = CharacEnumNum
                if str(sheet.cell(1,icol).value) == "STRING":
                    continue
                if str(sheet.cell(1,icol).value) == "TEXT":
                    continue
                for irow in range(5,sheet.nrows):
                    if str(sheet.cell(irow,0).value) == "":
                        continue
                    for c in str(sheet.cell(irow,icol).value):
                        if c not in CharList:
                            LogError(StrESRC(ExcelName,SheetName,irow+1,icol+1)+str(sheet.cell(irow,icol).value)+" Character not in "+str(CharList))
                            return False

            #ID列检查
            if(str(sheet.cell(1,1).value) != 'INT'):
                LogError(StrESRC(ExcelName,SheetName,2,2)+str(sheet.cell(1,1).value)+ " first colum is not INT!" )
                return False

            #变量名检查
            row_2 = sheet.row(2)
            for isign in range(1,len(row_2)):
                arg = str(row_2[isign].value)
                for c in arg:
                    if c not in CharacterList and str(sheet.cell(0,isign).value) != "":
                        LogError(StrESRC(ExcelName,SheetName,3,isign+1)+ arg + " not in a-z、A-Z、_、1-0")
                        return False

            #注释名检查
            row_4 = sheet.row(4)
            for isign in range(1,len(row_4)):
                cname = str(row_4[isign].value)
                if "\n" in cname[0:-1]:
                    LogError(StrESRC(ExcelName,SheetName,5,isign+1)+"content not allow \\n !")
                    return False

            #ID合法性
            col_1 = sheet.col(1)
            id_list = []
            for isign in range(5,len(col_1)):
                if str(sheet.cell(isign,0).value) == "":
                    continue
                str_id = str(col_1[isign].value)
                #print(str_id)
                if str_id == "":
                    LogError(StrESRC(ExcelName,SheetName,isign+1,2)+ "id is empty!")
                    return False
                if str_id in id_list:
                    LogError(StrESRC(ExcelName,SheetName,isign+1,2)+ "id is repeated!")
                    return False
                id_list.append(str_id)
            
            #关键标记检查
            if "*" not in str(sheet.cell(0,1).value):
                LogError(StrESRC(ExcelName,SheetName,2,1) + " need *|#*|*#")
                return False
            if "*" not in str(sheet.cell(1,0).value):
                LogError(StrESRC(ExcelName,SheetName,2,1) + " need *|#*|*#")
                return False
            if "*" not in str(sheet.cell(2,0).value):
                LogError(StrESRC(ExcelName,SheetName,3,1) + " need *|#*|*#")
                return False
    return True

def RunWrap(ret,log):
    if ret:
        LogInfo(log+" succeed!")
    else:
        LogWarning(log+" failed!")
    return ret

def ExcelPreCheck():
    if not RunWrap(GetExcelFiles(),"[1/9]GetExcelFiles"):return False
    if not RunWrap(CheckFileType(),"[2/9]CheckFileType"):return False
    if not RunWrap(CheckFileName(),"[3/9]CheckFileName"):return False
    if not RunWrap(CheckLoadXlsm(),"[4/9]CheckLoadXlsm"):return False
    if not RunWrap(CheckShtIndex(),"[5/9]CheckShtIndex"):return False
    if not RunWrap(CheckIndexFmt(),"[6/9]CheckIndexFmt"):return False
    if not RunWrap(CheckIndexCnt(),"[7/9]CheckIndexCnt"):return False
    if not RunWrap(CheckSheetNms(),"[8/9]CheckSheetNms"):return False
    if not RunWrap(CheckSheetFmt(),"[9/9]CheckSheetFmt"):return False

    LogFinal("Excel pre check succeed!")
    time.sleep(1)

    return True