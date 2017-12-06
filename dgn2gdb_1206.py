#-*-coding:utf-8-*-
# 根據xls參數控制來源資料是否需匯入目的資料檔中
# 資料需求-------------------------#source來源資料夾# target目的資料夾#mapping.xls 對應參數
# 來源資料格式為dgn,目的資料夾格式為gdb 其他檔案格式亦可
# 工作流程: 批次讀檔, 將xls內容存成陣列 -- > 用對應參數篩選來源資料 --> 將資料匯至目的資料夾檔案中
# 使用工具: arcpy(Exists, MakeFeatureLayer, SelectLayerByAttribute, CalculateField, Delete)

# --------------------------------(O)--------------------------------
import arcpy
import os
import xdrlib
import xlrd
import sys

# ----------------------------路徑設定-----------------------------
Path = "D:\\Ashley\\Github\\dgn2gdb\\"
source_path = Path + "source\\" # 來源資料夾
target_path =  Path + "target\\VMap2.gdb\\VMap2TLM\\" # 目的資料夾
data = xlrd.open_workbook(Path + "mapping.xls")
schema_type = "NO_TEST" # append 參數
field_mapping = "" # append 參數


# --------------------------------(A)--------------------------------
# 讀mapping.xls, 將使用到的欄存成陣列
table = data.sheet_by_name(u'Sheet1')
Layer = []
Color = []
Linetype = []
LineWt = []
FeatureType = []
FeatureClass = []
Subtype = []
Update = []

# i=0為mapping.xls中的欄位名稱
# 用.append將欄位內容存到陣列中
for i in range(1, table.nrows): 
     Layer.append(table.row(i)[4].value) #  str, 'Level *'
     Color.append(table.row(i)[5].value) #  int, 13
     Linetype.append(table.row(i)[6].value) #  str,'7' (Long dash-dot Line)
     LineWt.append(table.row(i)[7].value) #  int, 2
     FeatureType.append(table.row(i)[12].value) #  str, 'Polyline'
     FeatureClass.append(table.row(i)[13].value) #  str, 'PolbnL'
     Subtype.append(table.row(i)[15].value) #  int, 0
     Update.append(table.row(i)[16].value) # str, 'acc=1;use_=30;'

# fileList用於存放每一個.dgn路徑
fileList = []
for dirPath, dirNames, fileNames in os.walk(source_path):
     for f in fileNames:
          fileList.append(os.path.join(dirPath, f))

# --------------------------------(C)--------------------------------
#for k in range(0, int(len(fileList))-1): # 依序讀fileList中的項目(source_path)
for k in range(1, 9): # 依序讀fileList中的項目(source_path)

# --------------------------------(B)--------------------------------

# 依序讀mapping.xls每一列
# j=0為#(A)所存陣列的第一個字串
     print "k = " + str(k)
     print "test"
     for j in range(0, 5):
          feature = str(fileList[k]) + "\\"+ str(FeatureType[j]) # 確認是否有該屬性的物件
          print feature



# --------------------------------(D)--------------------------------
          if arcpy.Exists(feature):  # 若dgn中有該屬性物件，進行資料處理
               print "j = "+ str(j)
               sql_select = '"Layer" = \'{}\' AND "Linetype" = \'{}\' AND "Color" ={} AND "LineWt" ={} '.format(str(Layer[j]), str(Linetype[j]), int(Color[j]), int(LineWt[j]))  # 篩選資料sql
               print "sql_select : {}".format(sql_select)
               layer = arcpy.MakeFeatureLayer_management(feature) # 為篩選資料, 建立source 圖層
               selecting = arcpy.SelectLayerByAttribute_management(layer, "NEW_SELECTION",  sql_select)
               print "make layer success"
               counts_SelectSource = int(arcpy.GetCount_management(selecting).getOutput(0)) # 符合篩選條件的筆數
               lastsnumber = 0 - counts_SelectSource # 從末端選資料, 項次為負
              

               if counts_SelectSource > 0:  #若counts_SelectSource不為0, 則"有資料"須匯到相對應的圖徵
                    
                    target = target_path +  str(FeatureClass[j])
                    arcpy.Append_management(selecting, target, schema_type, field_mapping, int(Subtype[j])) # 將篩選出的資料貼到target

                    #------------------------更新資料
                    string = Update[j]
                    print Update[j]
                    if string.find('=') != 0:
                         layer2 = arcpy.MakeFeatureLayer_management(target)
                         rows = [row for row in arcpy.da.SearchCursor(layer2, "ObjectID")]
                         print rows
                         sql = "ObjectID >= {}".format(int(rows[lastsnumber][0]))
                         arcpy.SelectLayerByAttribute_management(layer2, "NEW_SELECTION",  sql)
                         for p in range (0, string.count(';')):
                              col = string.split(';')[p].split('=')[0]
                              val = string.split(';')[p].split('=')[1]
                              arcpy.CalculateField_management(layer2, col, val, "PYTHON", "")
                              print  "import " + str(FeatureType[j]) + " from " + str(fileNames[k]) + " to "+ FeatureClass[j] + ", " + "update info: " + string
                     #-------------------------更新資料

               
