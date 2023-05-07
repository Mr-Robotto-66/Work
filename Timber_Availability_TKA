# ---------------------------------------------------------------------------
# TimberAvailability.py
# Created on: May 5 , 2019
#
# Description: This tool is deisgned to find the available timber within each
#           operating area based on the constraints in the LUT_Processing table.
#           The script extracts VRI information for this available timber and
#           summarizes the data for the 6 leading species in each stand. It then
#           outputs the result to feature classes for each operating and maps the
#           result
#
# Output:   1 depleted feature class for the entire field team  based on the years
#           entered by the user in the tool. 1 feature class for each operating
#           of available mature timber (>80 years) and 1 feature class for each
#           operating area of approaching mature timber (60 to 80 years). PDF maps
#           for each operating area in the selected field team summarizing results
#
# Author:   Daniel Otto

#--------------------------------------------------------------------------------

#   Importing all modules that will be used for the script
from __future__ import division
import arcpy
from arcpy import env
import sys
import string
import os
import math
import stat
import time
import datetime
import getpass
import collections
from collections import OrderedDict

arcpy.env.overwriteOutput = True
arcpy.Delete_management("in_memory")

#   Set the workspace to be in memory so no extra data is created
arcpy.env.workspace = 'in_memory'

#   Set the output coordinate system to NAD 83 so any geoprocessing that occurs the
#   output is always in NAD 83.
arcpy.env.outputCoordinateSystem = sr = arcpy.SpatialReference(3005)

#  Pull arguements from the tool parameters window.
#--------------------------------------------------------------------------------

arcpy.AddMessage("Getting run parameters")
# Python script location (automatically populated)
PYTHON_SCRIPT = sys.argv[0]
#   User provides workspace in which to create new FGDB
file_path = sys.argv[1]
#   User identifies which field team they want to run Timber Availability on
FT = sys.argv[2]
# User enters the years to look back for depletion
deplete_years = int(sys.argv[3])
# User enters their name
author = sys.argv[4]

#   The Processing_Variables is the collection of messaging and script specific
#   information needed throughout this program.  It is constantly updated, and
#   generally includes all paths and configuration information.  For parameters
#   that are passed in to the script, they should generally be assigned to a
#   Processing_Variable through the Initialize() routine.
Processing_Variables = {}

def Initialize():

    arcpy.AddMessage("Initializing script...")
    #Define Variables
    Processing_Variables['outGDB'] = file_path + '\\' + FT + r'_TimberAvailabilityAnalysis.gdb'
    Processing_Variables['outGDBname'] = FT + '_TimberAvailabilityAnalysis.gdb'

    Processing_Variables['Script_Directory_Path'] = os.path.split(PYTHON_SCRIPT)[0]
    Processing_Variables['Supporting_Data_Directory_Path'] = Processing_Variables['Script_Directory_Path'] + r'\Supporting Data'
    Processing_Variables['Supporting_Data_GDB'] = Processing_Variables['Supporting_Data_Directory_Path'] + r'\SupportingData.gdb'

    Processing_Variables['ProcessingLookupTable'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_Processing'

    #Setup defintion query to look back user specdified years when determining depletion
    Processing_Variables['Deplete_year'] = str(int(datetime.datetime.now().year) - deplete_years)
    Processing_Variables['Deplete_exp'] = 'HARVEST_YEAR >= ' + Processing_Variables['Deplete_year']
    Processing_Variables['Deplete_layer'] = ''

    #Define select statement to select only operating areas from input field team
    Processing_Variables['selFT'] = "Field_Team = '" + FT + "'"

    #Create folder to put PDF maps
    arcpy.CreateFolder_management(file_path, "PDFs")
    Processing_Variables['Pdf_folder'] = file_path + r'\PDFs'

    #-------------------------------------------------------------------------------
    #  Script Controls
    #-------------------------------------------------------------------------------
    #This section creates a list of all the controls found in the LUT_ScriptControls
    #table and converts them all to variables.

    Processing_Variables['LUT_ScriptControls'] = Processing_Variables['Supporting_Data_GDB'] + '\\LUT_ScriptControls'

    Processing_Variables['Variables'] = {}
    for row in arcpy.da.SearchCursor(Processing_Variables['LUT_ScriptControls'],['Script_Variable','Variable_Value']):
        Processing_Variables['Variables'][row[0]] = row[1]

    #Set Operating Areas
    Processing_Variables['OpArea'] = Processing_Variables['Variables']['OpArea']
    Processing_Variables['OperatingAreas'] = ''
    Processing_Variables['OAnames'] = []

    # Set THLB
    Processing_Variables['THLB'] = Processing_Variables['Variables']['THLB']

    # Set VRI
    Processing_Variables['VRI'] = Processing_Variables['Variables']['VRI']

    # Set Depelted cutblocks
    Processing_Variables['CutBlk'] = Processing_Variables['Variables']['CutBlk']

    #Landscape and Portrait mxd templates
    Processing_Variables['Portrait_MXD'] = Processing_Variables['Variables']['Portrait_template']
    Processing_Variables['Landscape_MXD'] = Processing_Variables['Variables']['Landscape_template']

    #Operating Areas for mapping (all OA's formatted to fit on max 1:35000 map)
    Processing_Variables['OA_4_mapping'] = Processing_Variables['Variables']['OA_4_mapping']

    return(1)

def Setup():

    arcpy.AddMessage("Creating Geodatabase...")
    #Create File Geodatabase
    if arcpy.Exists(Processing_Variables['outGDB']):
        raise Exception('The FGDB, ' + Processing_Variables['outGDB'] + ' already exists! Rename any old versions prior to running this script.')
    else:
        print file_path
        arcpy.CreateFileGDB_management(file_path,Processing_Variables['outGDBname'])

    # Create a feature class in memeory to be used throughout the script
    arcpy.AddMessage("Selecting Operating areas from " + FT)
    arcpy.Select_analysis(Processing_Variables['OpArea'], r'in_memory\OperatingAreas', Processing_Variables['selFT'])
    Processing_Variables['OperatingAreas'] = r'in_memory\OperatingAreas'

    # Create a list of operating area names
    with arcpy.da.SearchCursor(Processing_Variables['OperatingAreas'],["OPERATING_AREA"]) as cursor:
        for row in cursor:
            Processing_Variables['OAnames'].append(row[0])

    # Clip depeleted cutblocks where harvest year was greater than 15 years ago
    #arcpy.MakeFeatureLayer_management(Processing_Variables['CutBlk'],"cutblk_FL")
    arcpy.AddMessage("Finding depleted blocks (harvested in last " + str(deplete_years) + " years) within Operating Areas")
    arcpy.Clip_analysis(Processing_Variables['CutBlk'],Processing_Variables['OperatingAreas'], r'in_memory\Cutblocks')
    arcpy.Select_analysis(r'in_memory\Cutblocks', r'in_memory\Deplete_layer', Processing_Variables['Deplete_exp'])
    Processing_Variables['Deplete_layer'] = r'in_memory\Deplete_layer'
    #Create a feature class of depletion for use in mapping
    arcpy.CopyFeatures_management(Processing_Variables['Deplete_layer'], file_path + '\\' + Processing_Variables['outGDBname'] + '\\Depleted')

    return(1)

def TimberAvailabilty():
    #Clip the THLB to the boundary of the operating areas
    arcpy.AddMessage("Clipping THLB to Operating areas: " + FT)
    arcpy.Clip_analysis(Processing_Variables['THLB'],Processing_Variables['OperatingAreas'], r'in_memory\THLB_OA')
    Processing_Variables['THLB_OA'] = r'in_memory\THLB_OA'
    arcpy.AddMessage("Finsished clipping THLB to Operating areas: " + FT)

    # Select only THLB with THLB_FACT > 0
    arcpy.AddMessage("Removing THLB polygons with THLB_FACT = 0 or NULL")
    arcpy.Select_analysis(Processing_Variables['THLB_OA'], r'in_memory\THLB_clipped', "THLB_FACT > 0 and TSA_NUMBER in ('11','18','15','23')")
    Processing_Variables['THLB_clipped'] = r'in_memory\THLB_clipped'

    # Remove depletion for the user defined number of years
    arcpy.AddMessage("Remvoing Depletion")
    arcpy.Erase_analysis(Processing_Variables['THLB_clipped'],Processing_Variables['Deplete_layer'], r'in_memory\THLB_deplete')
    Processing_Variables['THLB_clipped'] = r'in_memory\THLB_deplete'

    # Erase Items in LUT_Processing from THLB
    fieldnames = [f.name for f in arcpy.ListFields(Processing_Variables['ProcessingLookupTable']) if f.aliasName in ['Item',FT,'Source','Definition_Query']]
    for row in arcpy.da.SearchCursor(Processing_Variables['ProcessingLookupTable'], fieldnames):
        arcpy.AddMessage(str(row[0]) + str(row[1]) + str(row[2]) + str(row[3]))
        if row[1] > 0:
            #if no value in defintion query of layer remove whole layer
            if row[3] == None or row[3] =="":
                arcpy.Erase_analysis(Processing_Variables['THLB_clipped'], row[2], r'in_memory\THLB_erase'+str(row[1]))
                Processing_Variables['THLB_clipped'] = r'in_memory\THLB_erase'+str(row[1])
            #If there is a value in defintion query of layer set that before erase it
            else:
                arcpy.AddMessage("checck")
                def_query = "def_query"
                arcpy.MakeFeatureLayer_management(row[2], def_query)
                arcpy.SelectLayerByAttribute_management(def_query, "NEW_SELECTION", row[3])

                arcpy.AddMessage(str(row[3]))
                arcpy.Erase_analysis(Processing_Variables['THLB_clipped'], def_query, r'in_memory\THLB_erase'+str(row[1]))
                Processing_Variables['THLB_clipped'] = r'in_memory\THLB_erase'+str(row[1])
            arcpy.AddMessage(row[0] + ' removed from THLB')

    #Clip VRI to the the available THLB created above
    arcpy.AddMessage('Clipping VRI to Available THLB')
    arcpy.Select_analysis(Processing_Variables['VRI'], r'in_memory\VRI_select', "ORG_UNIT_CODE in ('DKA','DCS','DMH')")
    arcpy.Clip_analysis(r'in_memory\VRI_select',Processing_Variables['THLB_clipped'], r'in_memory\TA_VRI')
    Processing_Variables['TA_VRI'] = r'in_memory\TA_VRI'
    #intersect features in VRI layer so that operating areas which border each other are
    #separated, this will also add a operating field to the VRI layer which will be
    #used to separte OA's later
    intersect_features = [Processing_Variables['TA_VRI'],Processing_Variables['OperatingAreas']]
    arcpy.Intersect_analysis(intersect_features, r'in_memory\TA_VRI_OA')
    Processing_Variables['TA_VRI_OA'] = r'in_memory\TA_VRI_OA'
    arcpy.AddMessage('Completed Clipping VRI')

    return(1)

def VolumeCalculator():
    arcpy.AddMessage('Starting Volume Calculator')

    #Add fields to VRI FC
    arcpy.AddMessage('Adding fields to VRI FC')
    fieldGrpList = ['SPC1_GRP','SPC2_GRP','SPC3_GRP','SPC4_GRP','SPC5_GRP','SPC6_GRP']
    fieldVolList = ['HECTARES','SPC1_VOL_LIVE','SPC2_VOL_LIVE','SPC3_VOL_LIVE','SPC4_VOL_LIVE','SPC5_VOL_LIVE','SPC6_VOL_LIVE','SPC1_VOL_DEAD'
                    ,'SPC2_VOL_DEAD','SPC3_VOL_DEAD','SPC4_VOL_DEAD','SPC5_VOL_DEAD','SPC6_VOL_DEAD','SPC1_AREA','SPC2_AREA','SPC3_AREA','SPC4_AREA'
                    ,'SPC5_AREA','SPC6_AREA']
    speciesList = ['Pine', 'Fir', 'Hemlock', 'Larch', 'Cedar', 'Spruce', 'Balsam', 'Decid', 'Other']

    for grp in fieldGrpList:
        arcpy.AddField_management(Processing_Variables['TA_VRI_OA'], grp, "TEXT")
    for vol in fieldVolList:
        arcpy.AddField_management(Processing_Variables['TA_VRI_OA'], vol, "DOUBLE")
        #Calc all to zero instead of NULL so field can be used later in expression
        arcpy.CalculateField_management(Processing_Variables['TA_VRI_OA'],vol,0,"VB")
    for spc in speciesList:
        newField = spc + 'Vol'
        newField2 = spc + 'Area'
        arcpy.AddField_management(Processing_Variables['TA_VRI_OA'], newField, "DOUBLE")
        arcpy.AddField_management(Processing_Variables['TA_VRI_OA'], newField2, "DOUBLE")
        #Calc all to zero instead of NULL so field can be used later in expression
        arcpy.CalculateField_management(Processing_Variables['TA_VRI_OA'],newField,0,"VB")
        arcpy.CalculateField_management(Processing_Variables['TA_VRI_OA'],newField2,0,"VB")

    arcpy.AddField_management(Processing_Variables['TA_VRI_OA'], "PineVolDead", "DOUBLE")
    arcpy.CalculateField_management(Processing_Variables['TA_VRI_OA'],"PineVolDead",0,"VB")

    #Create lists for species classes
    arcpy.AddMessage('Creating lists for species class')
    pineList = ['PL','PLI','P','PW','PF','PA','PY','PJ']
    firList = ['FD','FDI']
    hemlockList = ['HW','H','HM']
    larchList = ['L','LA','LT','LW']
    cedarList = ['CW','C','YC']
    spruceList = ['SX','S','SS','SB','SE','SW']
    balsamList = ['BL','BA','B','BG']
    decidList = ['AC','AT','ACT','DR','E','EP','EA','MB']

    #Create and populate Species dictionary
    arcpy.AddMessage('Creating and populating Species dictionary')
    speciesDict = {}
    for x in pineList:
        speciesDict[x]= 'PINE'
    for x in firList:
        speciesDict[x]= 'FIR'
    for x in hemlockList:
        speciesDict[x]= 'HEMLOCK'
    for x in larchList:
        speciesDict[x]= 'LARCH'
    for x in cedarList:
        speciesDict[x]= 'CEDAR'
    for x in spruceList:
        speciesDict[x]= 'SPRUCE'
    for x in balsamList:
        speciesDict[x]= 'BALSAM'
    for x in decidList:
        speciesDict[x]= 'DECIDUOUS'

    #Calculate area and volumes for all species in each stand
    arcpy.AddMessage('Calculating Hectares field from GEOMETRY_area....')
    VRI_FL = 'VRI_FL'
    arcpy.MakeFeatureLayer_management(Processing_Variables['TA_VRI_OA'], VRI_FL)
    arcpy.CalculateField_management(VRI_FL, "HECTARES", "!SHAPE.area@hectares!", "PYTHON_9.3")
    for i in range(1,7):
        arcpy.AddMessage('Calculating area for species ' + str(i))
        areaHa = 'SPC' + str(i) + '_AREA'
        spcPct = 'SPECIES_PCT_' + str(i)
        selNull = '"' + spcPct + '" is NULL'
        #Select species percent fields that are NULL and calc to zeros
        arcpy.SelectLayerByAttribute_management(VRI_FL, "NEW_SELECTION", selNull)
        arcpy.CalculateField_management(VRI_FL, spcPct,0,"VB")
        arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION")
        #Calculate proportion of stand area based on species percent
        calcArea = "([HECTARES] * [" + spcPct + "]) / 100"
        arcpy.CalculateField_management(VRI_FL, areaHa, calcArea, "VB")

        arcpy.AddMessage('Calculating volume for species ' + str(i))
        specCd = 'SPECIES_CD_' + str(i)

        #Calculate volumes for Pine and Deciduous at lower utilization level
        selStatement = specCd + " in ('PL','PLI','AT','AC','ACT','E','ES','EP','MB','DR')"
        arcpy.SelectLayerByAttribute_management(VRI_FL, "NEW_SELECTION", selStatement)
        if arcpy.GetCount_management(VRI_FL) > 0:
            volHaLive = 'SPC' + str(i) + '_VOL_LIVE'
            volHaDead = 'SPC' + str(i) + '_VOL_DEAD'
            liveVolperHa = 'LIVE_VOL_PER_HA_SPP' + str(i) + '_125'
            deadVolperHa = 'DEAD_VOL_PER_HA_SPP' + str(i) + '_125'
            calcStatementLive = '[' + liveVolperHa + '] * [HECTARES]'
            calcStatementDead = '[' + deadVolperHa + '] * [HECTARES]'
            arcpy.AddMessage( 'Calculating' + volHaLive + 'for PL or Deciduous types')
            arcpy.CalculateField_management(VRI_FL, volHaLive , calcStatementLive , "VB")
            arcpy.CalculateField_management(VRI_FL, volHaDead , calcStatementDead , "VB")
            arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )
        else:
            arcpy.AddMessage( 'There are no Pine or Deciduous types in ' + specCd)
            arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )

        #Calculate volumes for all other species at higher utilization level
        selStatement2 = specCd + " not in ('PL','PLI','AT','ACT','AC','E','ES','EP','MB','DR')"
        arcpy.SelectLayerByAttribute_management(VRI_FL, "NEW_SELECTION", selStatement2 )
        if arcpy.GetCount_management(VRI_FL) > 0:
            volHaLive = 'SPC' + str(i) + '_VOL_LIVE'
            volHaDead = 'SPC' + str(i) + '_VOL_DEAD'
            liveVolperHa = 'LIVE_VOL_PER_HA_SPP' + str(i) + '_175'
            deadVolperHa = 'DEAD_VOL_PER_HA_SPP' + str(i) + '_175'
            calcStatementLive = '[' + liveVolperHa + '] * [HECTARES]'
            calcStatementDead = '[' + deadVolperHa + '] * [HECTARES]'
            arcpy.AddMessage( 'Calculating ' + volHaLive + ' for all other types')
            arcpy.CalculateField_management(VRI_FL, volHaLive , calcStatementLive , "VB")
            arcpy.CalculateField_management(VRI_FL, volHaDead , calcStatementDead , "VB")
            arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )
        else:
            arcpy.AddMessage('There are no non-Pine or non-deciduous types in ' + specCd)
            arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )

    #Use Species dictionary to calculate species group field and sum volume by species
    for s, c in speciesDict.iteritems():
        for i in range(1,7):
            arcpy.AddMessage( 'Creating species groups...')
            arcpy.AddMessage( 'Working on ' + s + ' = ' + c)
            specGrp = 'SPC' + str(i) + '_GRP'
            spcVol = 'SPC' + str(i) + '_VOL_LIVE'
            spcVolD = 'SPC' + str(i) + '_VOL_DEAD'
            spcArea = 'SPC' + str(i) + '_AREA'
            specCode = 'SPECIES_CD_' + str(i)
            selStatement = specCode + ' = ' + "'" + s + "'"
            spcType = '"' + c + '"'
            arcpy.AddMessage( selStatement)
            arcpy.AddMessage( spcType)
            arcpy.AddMessage( c)
            arcpy.SelectLayerByAttribute_management(VRI_FL, "NEW_SELECTION", selStatement)
            if arcpy.GetCount_management(VRI_FL) > 0:
                arcpy.CalculateField_management(VRI_FL, specGrp , spcType, "VB")
                #Sum volume by species into new field
                if c == "PINE":
                    print c == "PINE"
                    calcExpr = "[PineVol] + [" + spcVol + "]"
                    calcExpr2 = "[PineArea] + [" + spcArea + "]"
                    calcExpr3 = "[PineVolDead] + [" + spcVolD + "]"
                    print calcExpr
                    try:
                        arcpy.CalculateField_management(VRI_FL, "PineVol", calcExpr, "VB")
                        arcpy.CalculateField_management(VRI_FL, "PineArea", calcExpr2, "VB")
                        arcpy.CalculateField_management(VRI_FL, "PineVolDead", calcExpr3, "VB")
                        print 'Field Calculated!'
                    except:
                        print 'Calculate failed!'
                        raise
                elif c == "FIR":
                    calcExpr = "[FirVol] + [" + spcVol + "]"
                    calcExpr2 = "[FirArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "FirVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "FirArea", calcExpr2, "VB")
                elif c == "HEMLOCK":
                    calcExpr = "[HemlockVol] + [" + spcVol + "]"
                    calcExpr2 = "[HemlockArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "HemlockVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "HemlockArea", calcExpr2, "VB")
                elif c == "LARCH":
                    calcExpr = "[LarchVol] + [" + spcVol + "]"
                    calcExpr2 = "[LarchArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "LarchVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "LarchArea", calcExpr2, "VB")
                elif c == "CEDAR":
                    calcExpr = "[CedarVol] + [" + spcVol + "]"
                    calcExpr2 = "[CedarArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "CedarVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "CedarArea", calcExpr2, "VB")
                elif c == "SPRUCE":
                    calcExpr = "[SpruceVol] + [" + spcVol + "]"
                    calcExpr2 = "[SpruceArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "SpruceVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "SpruceArea", calcExpr2, "VB")
                elif c == "BALSAM":
                    calcExpr = "[BalsamVol] + [" + spcVol + "]"
                    calcExpr2 = "[BalsamArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "BalsamVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "BalsamArea", calcExpr2, "VB")
                elif c == "DECIDUOUS":
                    calcExpr = "[DecidVol] + [" + spcVol + "]"
                    calcExpr2 = "[DecidArea] + [" + spcArea + "]"
                    arcpy.CalculateField_management(VRI_FL, "DecidVol", calcExpr, "VB")
                    arcpy.CalculateField_management(VRI_FL, "DecidArea", calcExpr2, "VB")
            else:
                print specCode + ' has no values = ' + s

    arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )

    #Select remaining species and group as 'other' and calculate volume
    for i in range(1,7):
        specGrp = 'SPC' + str(i) + '_GRP'
        spcVol = 'SPC' + str(i) + '_VOL_LIVE'
        specCode = 'SPECIES_CD_' + str(i)
        spcType = "\"OTHER\""
        selStatement = specCode + ' IS NOT NULL AND (' + specGrp + ' IS NULL OR ' + specGrp + ' = \'' '\')'
        print selStatement
        arcpy.SelectLayerByAttribute_management(VRI_FL, "NEW_SELECTION", selStatement)
        if arcpy.GetCount_management(VRI_FL) > 0:
            arcpy.CalculateField_management(VRI_FL, specGrp , spcType, "VB")
            calcExpr = "OtherVol + " + spcVol + ""
            arcpy.CalculateField_management(VRI_FL, "OtherVol", calcExpr, "VB")
        else:
            print 'There are no NULL values in ' + specGrp

    arcpy.SelectLayerByAttribute_management(VRI_FL, "CLEAR_SELECTION" )

    #separate out Mature and Immature VRI for each OA
    for OA in arcpy.da.SearchCursor(Processing_Variables['OperatingAreas'],["SHAPE@","OPERATING_AREA"]):
        arcpy.AddMessage('Working on VRI analysis for ' + OA[1])
        OAname = OA[1].replace(" ","_")
        Processing_Variables['Immature_VRI_OA'] = r'in_memory\ImmatureVRI_OA'
        Processing_Variables['Mature_VRI_OA'] = r'in_memory\MatureVRI_OA'
        # Set Volume Exprressions
        Processing_Variables['ImmatureExp'] = "OPERATING_AREA = '" + OA[1] +"' and PROJ_AGE_1 > 60 and PROJ_AGE_1 <= 80"
        Processing_Variables['MatureExp'] = "OPERATING_AREA = '" + OA[1] +"' and PROJ_AGE_1 > 80"

        arcpy.Select_analysis(VRI_FL,Processing_Variables['Immature_VRI_OA'],Processing_Variables['ImmatureExp'])
        arcpy.Select_analysis(VRI_FL,Processing_Variables['Mature_VRI_OA'],Processing_Variables['MatureExp'])
        # Create FC for each operating area mature and immature VRI for mapping
        arcpy.CopyFeatures_management(Processing_Variables['Immature_VRI_OA'], file_path + '\\' + Processing_Variables['outGDBname'] + '\\Immature_VRI_' + OAname)
        arcpy.CopyFeatures_management(Processing_Variables['Mature_VRI_OA'], file_path + '\\' + Processing_Variables['outGDBname'] + '\\Mature_VRI_' + OAname)

        #start summary statistics
        if arcpy.Exists(r'in_memory\mature_stats'):
            arcpy.Delete_management(r'in_memory\mature_stats')
        if arcpy.Exists(r'in_memory\immature_stats'):
                    arcpy.Delete_management(r'in_memory\immature_stats')
        Processing_Variables['mature_stats'] = r'in_memory\mature_stats'
        Processing_Variables['immature_stats'] = r'in_memory\immature_stats'

        stats = [["HECTARES","SUM"],["PineArea","SUM"],["PineVol","SUM"],["FirArea","SUM"],["FirVol","SUM"],["HemlockArea","SUM"],["HemlockVol","SUM"],["LarchArea","SUM"],
                ["LarchVol","SUM"],["CedarArea","SUM"],["CedarVol","SUM"],["SpruceArea","SUM"],["SpruceVol","SUM"],["BalsamArea","SUM"],["BalsamVol","SUM"],["DecidArea","SUM"],
                ["DecidVol","SUM"],["OtherArea","SUM"],["OtherVol","SUM"]]
        arcpy.Statistics_analysis(Processing_Variables['Mature_VRI_OA'],Processing_Variables['mature_stats'],stats)
        arcpy.Statistics_analysis(Processing_Variables['Immature_VRI_OA'],Processing_Variables['immature_stats'],stats)

        #start mapping
        arcpy.Select_analysis(Processing_Variables['OA_4_mapping'], r'in_memory\OAmapping', "OPERATING_AREA = '" + OA[1] +"'" )
        for feature in arcpy.da.SearchCursor(r'in_memory\OAmapping',["SHAPE@","OA_4_mapping", "ORIENTATION", "SCALE"]):
            arcpy.AddMessage('Beginning the mapping for' + feature[1])
            if feature[2] == "P":
                mxd = arcpy.mapping.MapDocument(Processing_Variables['Portrait_MXD'])

            else:
                mxd = arcpy.mapping.MapDocument(Processing_Variables['Landscape_MXD'])

            df = arcpy.mapping.ListDataFrames(mxd)[0]
            lyr = arcpy.mapping.ListLayers(mxd, "OA_4_mapping")[0]

            #Select operating area
            selOA = "OA_4_mapping = '" + feature[1] + "'"
            arcpy.SelectLayerByAttribute_management(lyr,"NEW_SELECTION",selOA)

            #Zoom to selected opearting area
            arcpy.AddMessage("Zooming to operating area")
            df.extent = lyr.getSelectedExtent()
            df.scale = int(feature[3])

            #Update title of map to current operating area
            for elm in arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT"):
                if elm.name == 'Title':
                   elm.text = feature[1]
                if elm.name == 'file_path':
                   elm.text = 'Path: ' + Processing_Variables['Pdf_folder']
                if elm.name == 'author':
                   elm.text = 'Created by: ' + author
                #Update Mature Volume Summary table in MXD
                for ma_row in arcpy.da.SearchCursor(Processing_Variables['mature_stats'], ["SUM_HECTARES","SUM_PineArea","SUM_PineVol",
                                                    "SUM_FirArea","SUM_FirVol","SUM_HemlockArea","SUM_HemlockVol","SUM_LarchArea","SUM_LarchVol",
                                                    "SUM_CedarArea","SUM_CedarVol","SUM_SpruceArea","SUM_SpruceVol","SUM_BalsamArea","SUM_BalsamVol",
                                                    "SUM_DecidArea","SUM_DecidVol","SUM_OtherArea","SUM_OtherVol"]):
                    if elm.name == 'area_ma_total':
                        elm.text = str(int(ma_row[0]))
                    if elm.name == 'area_ma_pine':
                        elm.text = str(int(ma_row[1]))
                    if elm.name == 'vol_ma_pine':
                        elm.text = str(int(ma_row[2]))
                    if elm.name == 'area_ma_fir':
                        elm.text = str(int(ma_row[3]))
                    if elm.name == 'vol_ma_fir':
                        elm.text = str(int(ma_row[4]))
                    if elm.name == 'area_ma_hemlock':
                        elm.text = str(int(ma_row[5]))
                    if elm.name == 'vol_ma_hemlock':
                        elm.text = str(int(ma_row[6]))
                    if elm.name == 'area_ma_larch':
                        elm.text = str(int(ma_row[7]))
                    if elm.name == 'vol_ma_larch':
                        elm.text = str(int(ma_row[8]))
                    if elm.name == 'area_ma_cedar':
                        elm.text = str(int(ma_row[9]))
                    if elm.name == 'vol_ma_cedar':
                        elm.text = str(int(ma_row[10]))
                    if elm.name == 'area_ma_spruce':
                        elm.text = str(int(ma_row[11]))
                    if elm.name == 'vol_ma_spruce':
                        elm.text = str(int(ma_row[12]))
                    if elm.name == 'area_ma_balsam':
                        elm.text = str(int(ma_row[13]))
                    if elm.name == 'vol_ma_balsam':
                        elm.text = str(int(ma_row[14]))
                    if elm.name == 'area_ma_deciduous':
                        elm.text = str(int(ma_row[15]))
                    if elm.name == 'vol_ma_deciduous':
                        elm.text = str(int(ma_row[16]))
                    if elm.name == 'area_ma_other':
                        elm.text = str(int(ma_row[17]))
                    if elm.name == 'vol_ma_other':
                        elm.text = str(int(ma_row[18]))
                    if elm.name == 'vol_ma_total':
                        elm.text = str(int(ma_row[2] + ma_row[4] + ma_row[6] + ma_row[8] + ma_row[10] + ma_row[12] + ma_row[14] + ma_row[16] + ma_row[18]))

                #Update immature Volume Summary table in MXD
                for im_row in arcpy.da.SearchCursor(Processing_Variables['immature_stats'], ["SUM_HECTARES","SUM_PineArea","SUM_PineVol",
                                                    "SUM_FirArea","SUM_FirVol","SUM_HemlockArea","SUM_HemlockVol","SUM_LarchArea","SUM_LarchVol",
                                                    "SUM_CedarArea","SUM_CedarVol","SUM_SpruceArea","SUM_SpruceVol","SUM_BalsamArea","SUM_BalsamVol",
                                                    "SUM_DecidArea","SUM_DecidVol","SUM_OtherArea","SUM_OtherVol"]):
                    if elm.name == 'area_im_total':
                        elm.text = str(int(im_row[0]))
                    if elm.name == 'area_im_pine':
                        elm.text = str(int(im_row[1]))
                    if elm.name == 'vol_im_pine':
                        elm.text = str(int(im_row[2]))
                    if elm.name == 'area_im_fir':
                        elm.text = str(int(im_row[3]))
                    if elm.name == 'vol_im_fir':
                        elm.text = str(int(im_row[4]))
                    if elm.name == 'area_im_hemlock':
                        elm.text = str(int(im_row[5]))
                    if elm.name == 'vol_im_hemlock':
                        elm.text = str(int(im_row[6]))
                    if elm.name == 'area_im_larch':
                        elm.text = str(int(im_row[7]))
                    if elm.name == 'vol_im_larch':
                        elm.text = str(int(im_row[8]))
                    if elm.name == 'area_im_cedar':
                        elm.text = str(int(im_row[9]))
                    if elm.name == 'vol_im_cedar':
                        elm.text = str(int(im_row[10]))
                    if elm.name == 'area_im_spruce':
                        elm.text = str(int(im_row[11]))
                    if elm.name == 'vol_im_spruce':
                        elm.text = str(int(im_row[12]))
                    if elm.name == 'area_im_balsam':
                        elm.text = str(int(im_row[13]))
                    if elm.name == 'vol_im_balsam':
                        elm.text = str(int(im_row[14]))
                    if elm.name == 'area_im_deciduous':
                        elm.text = str(int(im_row[15]))
                    if elm.name == 'vol_im_deciduous':
                        elm.text = str(int(im_row[16]))
                    if elm.name == 'area_im_other':
                        elm.text = str(int(im_row[17]))
                    if elm.name == 'vol_im_other':
                        elm.text = str(int(im_row[18]))
                    if elm.name == 'vol_im_total':
                        elm.text = str(int(im_row[2] + im_row[4] + im_row[6] + im_row[8] + im_row[10] + im_row[12] + im_row[14] + im_row[16] + im_row[18]))


            #Clear selection
            arcpy.SelectLayerByAttribute_management(lyr,"CLEAR_SELECTION")

            Deplete_lyr = arcpy.mapping.ListLayers(mxd,"Depleted Blocks",df)[0]
            Deplete_lyr.replaceDataSource(Processing_Variables['outGDB'], "FILEGDB_WORKSPACE", 'Depleted')

            mature_lyr = arcpy.mapping.ListLayers(mxd,"Species - Mature (>80 years)",df)[0]
            mature_lyr.replaceDataSource(Processing_Variables['outGDB'], "FILEGDB_WORKSPACE", 'Mature_VRI_' + OAname)

            immature_lyr = arcpy.mapping.ListLayers(mxd,"Species - Approaching Mature (60-80 years)",df)[0]
            immature_lyr.replaceDataSource(Processing_Variables['outGDB'], "FILEGDB_WORKSPACE", 'Immature_VRI_' + OAname)

            arcpy.RefreshActiveView()

            #Export map to PDF and save
            arcpy.AddMessage("Exporting to PDF")
            outPDF = Processing_Variables['Pdf_folder'] + "\\" + feature[1] + "_Timber_Availability_" + str(int(df.scale)) + "K.pdf"
            arcpy.mapping.ExportToPDF(mxd,outPDF)

            arcpy.AddMessage("Map complete")


# Run Main for ArcGIS
#
#   Main Method manages the overall program
#---------------------------------------------------------------------------------------------------------
def runMain():

    Initialize()
    Setup()
    TimberAvailabilty()
    VolumeCalculator()

    return(1)

runMain()








