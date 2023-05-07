# ---------------------------------------------------------------------------
# BCTS_Data_Extractor_waste_management.py
# Created on: May 31, 2020
# Usage:
# Description:
# Output:
#
# Author: Daniel Otto
# ---------------------------------------------------------------------------
# Update:
#
# Author:
# ---------------------------------------------------------------------------

# Import modules
import arcpy, os, sys, string
arcpy.env.overwriteOutput = True

arcpy.Delete_management("in_memory")

#Script_Path
Processing_Variables = {}
PYTHON_SCRIPT = sys.argv[0]
Processing_Variables['Script_Directory_Path'] = os.path.split(PYTHON_SCRIPT)[0]

# Script arguments
#File Geodatabase in which to store output data
name = sys.argv[1]
outPath = sys.argv[2]
newFolder = name + "_Waste_Mgmt_Spatial"
Output_File_GDB = name + ".gdb"

# Set the qualified field names to false. Improtant so that field names match in both databases and correct fields will be populated in archive FC
arcpy.env.qualifiedFieldNames = False


#Expression to select a subset of the forest_13 block layer
select_type = sys.argv[3]
identifier = sys.argv[4].upper()
identifier = identifier.replace (" ", "")
identifier = identifier.replace (",", "','")

if select_type == "Licence":
    Expression = "LICENCE_ID in ('" + identifier + "')"
elif select_type == "Block":
    Expression = "BLOCK_ID in ('" + identifier + "')"
else:
    arcpy.AddMessage("Invalid Expression")

arcpy.AddMessage(Expression)

#User chooses if shapefile, gdb, KML output desired

exportgdb = sys.argv[5]
exportshp = sys.argv[6]
exportKML = sys.argv[7]
#Get username and password for mapview account
DBP06_username = sys.argv[8]
DBP06_password = sys.argv[9]

#Set workspace
arcpy.Delete_management("in_memory")
arcpy.env.workspace = "in_memory"

def Create_folder():
    # Create New Folder for output data
    arcpy.AddMessage("Create new output folder...")
    Processing_Variables['Output_folder'] = outPath + '\\' + newFolder
    if arcpy.Exists(Processing_Variables['Output_folder']):
        msg = "There is already a workspace called " + newFolder + "."
        print msg
        arcpy.AddMessage("----------------------------------------")
        arcpy.AddMessage(msg)
        arcpy.AddMessage("----------------------------------------")
        sys.exit()
    else:
        arcpy.AddMessage("Creating new output folder...")
        arcpy.CreateFolder_management(outPath,newFolder)

def Delete_fields(keep_list, input_layer): #delete unwated fields from a layer
    #split the users list
    keepFieldList = keep_list.split(",")
    # Use ListFields to get a list of fields in Block layer
    fieldObjList = arcpy.ListFields(input_layer)

    # Create an empty list that will be populated with field names
    fieldNameList = []
    # For each field in the object list, add the field name to the
    #  name list.  If the field is required, exclude it, to prevent errors
    for field in fieldObjList:
        if not field.required:
            fieldNameList.append(field.name)
##
##    arcpy.AddMessage(keepFieldList)
##    arcpy.AddMessage(fieldNameList)
    #Loop through list of fields to keep and remove them from the field name list
    for k in keepFieldList:
        fieldNameList.remove(k)



    #Delete unwated fields
    arcpy.DeleteField_management(input_layer, fieldNameList)

def DBP06(username,password,path): # Get and pass DBP06 map_view credtiantls
    arcpy.AddMessage('running DBP06 connection')
##    arcpy.AddMessage(username)
##    arcpy.AddMessage(password)
    name = 'Temp_DBP06'
    Processing_Variables['Temp_DBP06'] = path + r'\Temp_DBP06.sde'
    database_platform = 'ORACLE'
    account_authorization  = 'DATABASE_AUTH'
    instance = 'lrmbctsp.nrs.bcgov/dbp06.nrs.bcgov'

    if arcpy.Exists(os.path.join(path,'Temp_DBP06.sde')):
        try:
            os.remove(os.path.join(path,'Temp_DBP06.sde'))
        except:
            pass

    arcpy.CreateDatabaseConnection_management (path,name,database_platform,instance,account_authorization,username ,password,'DO_NOT_SAVE_USERNAME')

def Initialize():
    #Grab the variables from the script control table
    Processing_Variables['Supporting_Data_GDB'] = Processing_Variables['Script_Directory_Path'] + r'\Supporting_Data.gdb'
    Processing_Variables['ScriptControls'] = Processing_Variables['Supporting_Data_GDB'] + r'\Script_Controls_waste_management'

    Processing_Variables['Variables'] = {}
    for row in arcpy.da.SearchCursor(Processing_Variables['ScriptControls'],['Script_Variable','Variable_Value']):
        Processing_Variables['Variables'][row[0]] = row[1]

    Processing_Variables['Variable_Join'] = {}
    for row in arcpy.da.SearchCursor(Processing_Variables['ScriptControls'],['Script_Variable','Variable_table_join']):
        Processing_Variables['Variable_Join'][row[0]] = row[1]

    Processing_Variables['Variable_fields'] = {}
    for row in arcpy.da.SearchCursor(Processing_Variables['ScriptControls'],['Script_Variable','variable_field_list']):
        Processing_Variables['Variable_fields'][row[0]] = row[1]


    Processing_Variables['Blocks'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variables']['CUT_BLOCK_SHAPE']
    Processing_Variables['SU'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variables']['SU_SHAPE']
    Processing_Variables['SU_TABLE'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variable_Join']['SU_SHAPE']
    Processing_Variables['Harvest'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variables']['HARVEST_SHAPE']
    Processing_Variables['Falling_Corners'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variables']['FC_SHAPE']
    Processing_Variables['ROAD'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variables']['ROAD_SHAPE']
    Processing_Variables['ROAD_TABLE'] = Processing_Variables['Temp_DBP06'] + "\\" + Processing_Variables['Variable_Join']['ROAD_SHAPE']


def copy_features():
    #Select the blocks
    arcpy.Select_analysis(Processing_Variables['Blocks'],r'in_memory/blocks',Expression)
    #keep only the fields specified in the BLOCK_FIELD_LIST from the script controls table
    Delete_fields(Processing_Variables['Variable_fields']['CUT_BLOCK_SHAPE'],r'in_memory/blocks')

    #get the cutblock sequence numbers of the selected block
    Processing_Variables['SEQ_NUM_LIST'] = [str(row[0])[:-2] for row in arcpy.da.SearchCursor(r'in_memory/blocks',['CUTB_SEQ_NBR'])]
    Processing_Variables['SEQ_NUM_EXP'] = 'CUTB_SEQ_NBR in (' +','.join(Processing_Variables['SEQ_NUM_LIST']) + ')'
    arcpy.AddMessage(Processing_Variables['SEQ_NUM_EXP'])

    # Join SU table to SU layer
    arcpy.MakeFeatureLayer_management(Processing_Variables['SU'], "SU_FL")
    arcpy.AddJoin_management("SU_FL", "STUN_SEQ_NBR", Processing_Variables['SU_TABLE'], "STUN_SEQ_NBR")
    arcpy.Select_analysis("SU_FL", r'in_memory/SU', Processing_Variables['SEQ_NUM_EXP'])
    #keep only the fields specified in the SU_FIELD_LIST from the script controls table
    Delete_fields(Processing_Variables['Variable_fields']['SU_SHAPE'],r'in_memory/SU')

    #Get the harvest unit shapes
    arcpy.Select_analysis(Processing_Variables['Harvest'], r'in_memory/Harvest', Processing_Variables['SEQ_NUM_EXP'])
    Delete_fields(Processing_Variables['Variable_fields']['HARVEST_SHAPE'],r'in_memory/Harvest')

    #Get the Falling corners for selected blocks
    Processing_Variables['FC_EXP'] = Processing_Variables['SEQ_NUM_EXP'] + " and HUB_TYPE = 'FC'"
    arcpy.Select_analysis(Processing_Variables['Falling_Corners'], r'in_memory/FC', Processing_Variables['FC_EXP'])
    Delete_fields(Processing_Variables['Variable_fields']['FC_SHAPE'],r'in_memory/FC')

    #Get the roads within 100m of selected blocks
    arcpy.MakeFeatureLayer_management(Processing_Variables['ROAD'], "ROAD_FL")
    arcpy.MakeRouteEventLayer_lr("ROAD_FL","ROAD_SEQ_NBR", Processing_Variables['ROAD_TABLE'], "ROAD_SEQ_NBR LINE RSTA_START_METRE_NBR RSTA_END_METRE_NBR","ROAD_EV_FL")
    arcpy.SelectLayerByLocation_management("ROAD_EV_FL", "INTERSECT", r'in_memory/blocks', Processing_Variables['Variables']['ROAD_INT_DIST'], "NEW_SELECTION")
    arcpy.CopyFeatures_management("ROAD_EV_FL",r'in_memory/roads')
    Delete_fields(Processing_Variables['Variable_fields']['ROAD_SHAPE'],r'in_memory/roads')


def copy_2_FC():
    #Create FGDB
    arcpy.CreateFileGDB_management(Processing_Variables['Output_folder'],"Waste_Mgmt_Spatial.gdb","Current")
    Processing_Variables['Output_gdb'] = Processing_Variables['Output_folder'] +"\\Waste_Mgmt_Spatial.gdb"

    #Copy features if there are more than 0 features in layer
    arcpy.CopyFeatures_management(r'in_memory/blocks',Processing_Variables['Output_gdb'] + r'\Block')

    if int(arcpy.GetCount_management(r'in_memory/SU').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/SU',Processing_Variables['Output_gdb'] + r'\SU')

    if int(arcpy.GetCount_management(r'in_memory/Harvest').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/Harvest',Processing_Variables['Output_gdb'] + r'\Harvest_Unit')

    if int(arcpy.GetCount_management(r'in_memory/FC').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/FC',Processing_Variables['Output_gdb'] + r'\Falling_Corners')

    if int(arcpy.GetCount_management(r'in_memory/roads').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/roads',Processing_Variables['Output_gdb'] + r'\Roads')

def copy_2_SHP():
    #Copy features if there are more than 0 features in layer
    arcpy.CreateFolder_management(Processing_Variables['Output_folder'],"Shapefiles")
    Processing_Variables['Shapefile_folder'] = Processing_Variables['Output_folder'] + '\\Shapefiles'
    arcpy.CopyFeatures_management(r'in_memory/blocks',Processing_Variables['Shapefile_folder'] + r'\Block.shp')

    if int(arcpy.GetCount_management(r'in_memory/SU').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/SU',Processing_Variables['Shapefile_folder'] + r'\SU.shp')

    if int(arcpy.GetCount_management(r'in_memory/Harvest').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/Harvest',Processing_Variables['Shapefile_folder'] + r'\Harvest_Unit.shp')

    if int(arcpy.GetCount_management(r'in_memory/FC').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/FC',Processing_Variables['Shapefile_folder'] + r'\Falling_Corners.shp')

    if int(arcpy.GetCount_management(r'in_memory/roads').getOutput(0)) > 0:
        arcpy.CopyFeatures_management(r'in_memory/roads',Processing_Variables['Shapefile_folder'] + r'\Roads.shp')

def copy_2_KML():
    #Add layers to mxd, apply desired symbology and export map to one kmz file
    #SAVE MXD PROVIDED AND CHANGE PATH TO LOCATION SAVED BELOW
    mxd = arcpy.mapping.MapDocument(Processing_Variables['Script_Directory_Path'] + r"\Waste_Mgmt_Data_Template.mxd")
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]

    #Replace data source of layers in mxd template
    for lyr in arcpy.mapping.ListLayers(mxd):
        if lyr.name == "Blocks":
            lyr.replaceDataSource(Processing_Variables['Output_gdb'],'FILEGDB_WORKSPACE',"Block")

        if lyr.name == "SUs":
            if int(arcpy.GetCount_management(r'in_memory/SU').getOutput(0)) > 0:
                lyr.replaceDataSource(Processing_Variables['Output_gdb'],'FILEGDB_WORKSPACE',"SU")
            else:
                arcpy.mapping.RemoveLayer(df, lyr)

        if lyr.name == "Harvest Units":
            if int(arcpy.GetCount_management(r'in_memory/Harvest').getOutput(0)) > 0:
                lyr.replaceDataSource(Processing_Variables['Output_gdb'],'FILEGDB_WORKSPACE',"Harvest_Unit")
            else:
                arcpy.mapping.RemoveLayer(df, lyr)

        if lyr.name == "Falling Corners":
            if int(arcpy.GetCount_management(r'in_memory/FC').getOutput(0)) > 0:
                lyr.replaceDataSource(Processing_Variables['Output_gdb'],'FILEGDB_WORKSPACE',"Falling_Corners")
            else:
                arcpy.mapping.RemoveLayer(df, lyr)

        if lyr.name == "Roads":
            if int(arcpy.GetCount_management(r'in_memory/roads').getOutput(0)) > 0:
                lyr.replaceDataSource(Processing_Variables['Output_gdb'],'FILEGDB_WORKSPACE',"Roads")
            else:
                arcpy.mapping.RemoveLayer(df, lyr)
    Processing_Variables['Output_MXD'] = Processing_Variables['Output_folder'] + '\\Waste_Mgmt_Referral_MXD.mxd'
    mxd.saveACopy(Processing_Variables['Output_MXD'])
    arcpy.MapToKML_conversion(Processing_Variables['Output_MXD'],"Layers", Processing_Variables['Output_folder'] +"\\"+ name +".kmz", 20000)
    arcpy.Delete_management(Processing_Variables['Output_MXD'])

if __name__ == '__main__':
    Create_folder()
    DBP06(DBP06_username,DBP06_password,Processing_Variables['Script_Directory_Path'])
    Initialize()
    copy_features()
    copy_2_FC()
    if exportshp == 'true':
        copy_2_SHP()
    if exportKML == 'true':
        copy_2_KML()
    if exportgdb == 'false':
        arcpy.Delete_management(Processing_Variables['Output_gdb'])

    os.remove(Processing_Variables['Temp_DBP06'])

