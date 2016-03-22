# -*- coding:utf-8 -*-

import string, os, sys, xdrlib, xlrd, struct
from xlrd import open_workbook
from StreamOp import StreamOp


VERSION_BIN = 1
BIN_EXCEL_EXT = '.bytes'

def getExcelFileList(dir):
    try:
        files = os.listdir(dir)
    except Exception as e:
        print (str(e))

    excels = []
    for f in files:
        if f.startswith('.'):
            continue
        if f.lower().endswith('.xls') or f.lower().endswith('.xlsx'):
            excels.append(dir + os.sep + f)

    excels.sort()
    return excels

def open_excel(path):
    try:
        data = xlrd.open_workbook(path)
        return data
    except Exception as e:
        print (str(e))

def getDataTypeValue(dataType):
    result = 0
    if dataType == 'int' or dataType == 'int32':
        result = 1
    elif dataType == 'string':
        result = 2
    elif dataType == 'bool':
        result = 3
    elif dataType == 'float':
        result = 4
    elif dataType == 'int16':
        result = 5
    elif dataType == 'int8':
        result = 6

    return result

def serializeHead(stream, convertColumnCount, keyCol, needConvert, columnsName, columnsType, onlyForClient):
    if keyCol > convertColumnCount:
        print ('key列超出最大列数')
        return False
    StreamOp.WriteInt(stream, VERSION_BIN)
    StreamOp.WriteInt16(stream, convertColumnCount)
    actualKeyCol = keyCol
    if onlyForClient == True:
        actualKeyCol = 0
        for i in range(len(needConvert)):
            if needConvert[i] == False:
                continue
            if keyCol == i:
                break
            actualKeyCol += 1
    StreamOp.WriteInt16(stream, actualKeyCol)

    for i in range(len(needConvert)):
        if needConvert[i] == False:
            continue
        StreamOp.WriteString(stream, columnsName[i])
        typeNum = getDataTypeValue(columnsType[i])
        if typeNum == 0:
            print (str(i+1)+'列类型错误，无法识别')
            return False
        StreamOp.WriteInt8(stream, typeNum)
        
    return True

def serializeContent(stream, dtype, content):
    #print ('serializeContent ' + str(content) + ' ' + str(stream.tell()))
    if dtype == 'int' or dtype == 'int32':
        if len(str(content)) == 0:
            content = 0
        StreamOp.WriteInt(stream, int(content))
    elif dtype == 'string':
        StreamOp.WriteString(stream, (str(content)).decode("utf-8"))
    elif dtype == 'bool':
        content = content.strip();
        if len(content) == 0 or content == '0' or content == '0.0' or content.lower() == 'false' or content == '假':
            StreamOp.WriteBool(stream, False)
        else:
            StreamOp.WriteBool(stream, True)
    elif dtype == 'float':
        if len(str(content)) == 0:
            content = 0
        StreamOp.WriteFloat(stream, float(content))
    elif dtype == 'int16':
        if len(str(content)) == 0:
            content = 0
        StreamOp.WriteInt16(stream, content)
    elif dtype == 'int8':
        if len(str(content)) == 0:
            content = 0
        StreamOp.WriteInt8(stream, content)

def getValue(value):
    if isinstance(value,unicode):
        return value.encode("utf-8")
    elif isinstance(value,int):
        return int(value)
    else:
        return value

def convertFile(srcPath, destDir, onlyForClient):

    print ('开始转换 ' + srcPath)
    data = open_excel(srcPath)
    try:
        table = data.sheet_by_name('output')
    except Exception:
        print ("no sheet named output")
        return

    name = os.path.basename(srcPath)
    name = os.path.splitext(name)[0]
    destPath = destDir + os.sep + name + BIN_EXCEL_EXT
    file = open(destPath, 'wb')
    
    rowCount = table.nrows
    columnCount = table.ncols

    #print ('rowCount=' + str(rowCount))
    #print ('columnCount=' + str(columnCount))

    keyCol = -1
    convertColumnCount = 0
    needConvert = []
    columnName = []
    columnType = []

    if onlyForClient == True:
        for i in range(columnCount):
            needConvert.append(False)
    else:
        convertColumnCount = columnCount
        for i in range(columnCount):
            needConvert.append(True)

    rowIdx = -1
    colIdx = 0
    content = ''

    for irow in range(rowCount):
        colIdx = 0
        row = table.row_values(irow)
        #print (row)
        for icol in range(columnCount):
            
            #print (row[icol])
            #content = str(row[icol])
            content = getValue(row[icol])
            #print (content)
            
            #print (str(irow)+'行'+str(icol)+'列 '+content)

            # 如果第一列是以双斜线开头，则为注释行，忽略
            if colIdx == 0 and isinstance(content, unicode) and content.startswith('//'):
                break

            # 如果第一列是以#开头，则为有效起始行，
            # !表示该列是key，默认第一列为key
            # *表示该列客户端要用到
            if rowIdx < 0 and colIdx == 0 and content.startswith('#'):
                rowIdx = 0
            if rowIdx < 0:
                break

            if rowIdx == 0:
                if content.find('!') >= 0:
                    if keyCol >= 0:
                        print ('重复设置key: ' + str(colIdx + 1) + '列')
                        return False
                    
                    keyCol = colIdx

                if content.find('k') >= 0 and onlyForClient == True and content.find('*') >= 0:
                    if keyCol >= 0:
                        print ('重复设置客户端key: ' + str(colIdx + 1) + '列')
                        #return False
                    
                    keyCol = colIdx
                    
                if onlyForClient == True and content.find('*') >= 0:
                    needConvert[colIdx] = True
                    #print (content + str(colIdx))
                    convertColumnCount += 1
                    
            elif rowIdx == 1:
                tempName = content.strip()
                if len(tempName) == 0:
                    print (str(colIdx + 1) + '列名称不能为空')
                    return False
                
                for ic in range(0, colIdx):
                    if columnName[ic] == tempName:
                        print ('第'+str(colIdx+1)+'列和第'+str(ic+1)+'列名称重复了')
                        return False
                
                #columnName[colIdx] = tempName
                columnName.insert(colIdx, tempName)
                
            elif rowIdx == 2:
                tempName = content.strip()
                if len(tempName) == 0:
                    print (columnName[colIdx] + '列类型不能为空')
                    return False
                
                #columnType[colIdx] = tempName
                columnType.insert(colIdx, tempName)
                
            else:
                if colIdx == keyCol and columnType[colIdx] == 'string':
                    content = content.strip()
                if colIdx == keyCol and (content == 0 or len(str(content)) == 0) and columnType[colIdx] == 'string':
                    print ('key为空，跳过第%r行' % rowIdx)
                    break;
                #print ('serializeContent ' + str(colIdx))
                if needConvert[colIdx] == True:
                    serializeContent(file, columnType[colIdx], content)

            colIdx += 1

        if rowIdx == 2:
            if keyCol < 0:
                keyCol = 0
            if columnType[keyCol].find('int') < 0 and columnType[keyCol] != 'string':
                print ('主key只能是int或string')
                return False;
            if serializeHead(file, convertColumnCount, keyCol, needConvert, columnName, columnType, onlyForClient) == False:
                return False
            
        if rowIdx >= 0:
            rowIdx += 1
    file.close()
    print (srcPath + '转换成功！')

def main():
    #for i in range(len(sys.argv)):
    #    print ("第%d个参数是：%s" % (i, sys.argv[i]))
    if len(sys.argv) < 4:
        print ("参数不能少于3个：excel路径 bin路径 forClient/forServer")
        return

    excelDir = sys.argv[1]

    pathOutput = sys.argv[2]
	
    onlyForClient = False
    if sys.argv[3] == "forClient":
        onlyForClient = True

    if os.path.exists(pathOutput) == False:
        os.mkdir(pathOutput)
        
    excels = getExcelFileList(excelDir)
    for f in excels:
        #print (f)
        convertFile(f, pathOutput, onlyForClient)

if __name__=="__main__":
    main()
