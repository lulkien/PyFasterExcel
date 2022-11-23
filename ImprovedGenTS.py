from ReadXlsxClass import ReadXlsx
from TsWriterClass import TsWriter
from openpyxl import load_workbook
from os import path
from time import time
import sys

surfix = {
    "KOR"             :  "ko_KR", 
    "US_ENG"          :  "en_US", 
    "US_FRENCH"       :  "fr_CA", 
    "US_SPAIN"        :  "es_US", 
    "US_PORTUGUESE"   :  "pt_BR", 
    "CHINA"           :  "zh_CN", 
    "ARABIC"          :  "ar_SA", 
    "UK_ENG"          :  "en_GB", 
    "PORTUGUESE"      :  "pt_PT", 
    "SPAIN"           :  "es_ES", 
    "FRENCH"          :  "fr_FR", 
    "ITALIAN"         :  "it_IT", 
    "GERMAN"          :  "de_DE", 
    "RUSSIAN"         :  "ru_RU", 
    "DUTCH"           :  "nl_NL", 
    "SWEDISH"         :  "sv_SE", 
    "POLISH"          :  "pl_PL", 
    "TURKISH"         :  "tr_TR", 
    "CZECH"           :  "cs_CZ", 
    "DANISH"          :  "da_DK", 
    "SLOVAKIA"        :  "sk_SK", 
    "NORWEGIAN"       :  "no_NO", 
    "HUNGARIAN"       :  "hu_HU", 
    "JP"              :  "jp_JP", 
    "AU ENG"          :  "en_AU", 
}

start_time = time() * 1000

if len(sys.argv) <= 1:
    raise SyntaxError("Need a xlsx file")

xlsxName = sys.argv[1]
if not path.exists(xlsxName):
    raise FileNotFoundError("{file} not found".format(file=xlsxName))
if not xlsxName.endswith(".xlsx"):
    raise NameError("It's not .xlsx file")

reader = ReadXlsx(xlsxName)
reader.setWorksheetData(0) 

def writeTsFile(fileName, locale_id, category, id_index, data_index) -> None:
    print("Write data to file {name}".format(name=fileName))
    writer = TsWriter(fileName)
    writer.writeHeader()
    writer.addAttributes(["version", "language"], ["2.0", locale_id])
    writer.openTagWithAttribute("TS")
    writer.openTag("context")
    writer.openInlineTag("name", category)

    for index in range(1, reader.totalRows()):
        idValue = reader.getValueAt(index, id_index).value
        dataValue = reader.getValueAt(index, data_index).value
        if not idValue:
            print("Break due to end of file")
            break
        if not dataValue:
            continue
        writer.openTag("message")
        writer.openInlineTag("source", str(idValue).strip())
        writer.openInlineTag("translation", str(dataValue).strip())
        writer.closeTag()   # close tag writer

    writer.closeTag()   # close tag context
    writer.closeTag()   # close tag TS

def main():
    print("GenTS-tool version: 1.4")
    # init data
    title_row = reader.getItemsAtRow(0)
    list_keys = list(surfix.keys())

    appName = "X_X_X_X"
    category = "X_X_X_X"
    idColumn = -1

    # pre-process
    for index in range(0, reader.totalColumns()):
        cell_value = title_row[index].value
        if str(cell_value).strip() == "APP":
            appName = reader.getValueAt(1, index).value
        elif str(cell_value).strip() == "CATEGORY":
            category = reader.getValueAt(1, index).value
        elif str(cell_value).strip() == "STR_ID":
            idColumn = index
        elif not cell_value:
            #print("Break due to empty")
            break

    # Validate file
    if appName == "X_X_X_X" or category == "X_X_X_X" or idColumn == -1:
        raise ImportError("It's not master language files")
    
    print("===============================================")
    print("Total rows: " + str(reader.totalRows()))
    print("Total columns: " + str(reader.totalColumns()))
    print("Application name: " + str(appName))
    print("Category: " + str(category))
    print("STR_ID column: " + str(idColumn))
    print("===============================================")

    # main-process
    for index in range(0, reader.totalColumns()):
        cell_value = title_row[index].value
        if cell_value in list_keys:
            locale_id = surfix[cell_value]
            toFile = "output/{app}/{app}_{sf}.ts".format(app=appName, sf=locale_id)
            writeTsFile(toFile, locale_id, category, idColumn, index)
        elif not cell_value:
            print("Break due to empty")
            break


if __name__ == "__main__":
    main()

finish_time = time() * 1000
print("===============================================")
print("Elapsed time: " + str(finish_time - start_time) + "ms")