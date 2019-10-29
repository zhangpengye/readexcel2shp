#coding=utf-8
import arcpy
import xlrd
import os
import sys
from arcpy import env
import codecs
reload(sys)
sys.setdefaultencoding('utf-8')
path = r'E:\Python27\ArcGIS10.2\gx'
filenames = os.listdir(path.decode('utf-8'))
#获取shape文件保存的地址
#outWorkspace = os.path.split(outfile)[0]
#print(outWorkspace)
#shape文件保存的名字
#shpName = os.path.split(outfile)[-1]
#gp.CreateFeatureClass_management(outWorkspace,shpName,featureType,"","","",)
for filename in filenames:
    if ".xlsx" in filename:
        #print(filename)
        fileNm = os.path.splitext(filename)[0]
        #print(fileNm)
        goal_path = path+'\\'+fileNm
        os.mkdir(goal_path)
        print(goal_path)
        #outfile = path + "\\"+fileNm+'.shp'
        #print(outfile)
        #获取shape文件保存的地址
        #outWorkspace = os.path.split(outfile)[0]
        #print(outWorkspace)
        #shape文件保存的名字
        #shpName = os.path.split(outfile)[-1]
        #print(shpName)
        shpName = fileNm+'.shp'
        print(shpName)
        point = arcpy.Point()
        spRef = arcpy.SpatialReference(4490)
        #print(shpName)
        #gp = arcpy.create()
        #gp.OverwriteOutput = 1
        env.workspace = "C:/data"
        arcpy.CreateFeatureclass_management(goal_path,shpName,"POINT","","","",spRef)
        outfile = goal_path+"\\"+shpName
        print(outfile)
        arcpy.AddField_management(outfile,"name","text","","","100")
        arcpy.AddField_management(outfile,"typecode","text","","","100")
        arcpy.AddField_management(outfile,"address","text","","","250")
        arcpy.AddField_management(outfile,"type","text","","","100")
        arcpy.AddField_management(outfile,"tel","text","","","100")
        data = xlrd.open_workbook(filename)
        table = data.sheets()[0]
        nrows = table.nrows
        ncols = table.ncols
        print(nrows)
        #pointArray = []
        pointGeometryList = []
        #获取插入游标
        cur = arcpy.da.InsertCursor(outfile,("SHAPE@XY","name","typecode","address","type","tel"))
        #with arcpy.InsertCursor(outfile,("SHAPE@XY","name"))as cur:
         #   cntr =1
        #i = 1
        for i in range(1,nrows):
            #
            cell_X = table.cell(i,5).value
            cell_Y = table.cell(i,6).value
            #point.X = float(cell_X)
            lon = float(cell_X)
            cell_Name = table.cell(i,0).value
            cell_Address = table.cell(i,2).value
            cell_Typecode = table.cell(i,1).value
            cell_Tel = table.cell(i,3).value
            cell_Type = table.cell(i,4).value
            #newRow.setValue("name",cell_Name)
            #newRow.setValue()
            #point.Y = float(cell_Y)
            lat = float(cell_Y)
            #row_value = [(lon,lat),cell_Name]
            #rowValue = [(lon,lat),cell_Name]
            
            #pointGeometry = arcpy.PointGeometry(point,spRef)
                #pointGeometryList.append(pointGeometry)
            #newRow.shape = pointGeometry
            cur.insertRow(((lon,lat),cell_Name,cell_Typecode,cell_Address,cell_Type,cell_Tel))
                #newRow.name = table.cell(i,0)
                #cur.InsertRow(newRow)
            #print(table.cell(i,0))
           

            