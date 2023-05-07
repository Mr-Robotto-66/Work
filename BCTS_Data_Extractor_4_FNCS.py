# ---------------------------------------------------------------------------
# BCTS_Data_Extractor_4_FNCS.py
# Created on: Feb 23, 2023
# Usage: This script extracts forest blocks from a database and exports them to a file geodatabase,
#        a shapefile, and/or a KML file. The script takes user inputs for selecting the desired
#        forest blocks, the desired output types (file geodatabase, shapefile, and/or KML file),
#        and the output location. The script also performs some data processing tasks, including
#        selecting and merging the desired forest blocks, dissolving them into a single polygon,
#        and deleting unnecessary fields.
#
# Description: This script is intended to be used in the British Columbia Timber Sales (BCTS) program
#              for extracting forest blocks from a database and exporting them for use in the First
#              Nation Consultation System (FNCS).
#
# Output: The script outputs the selected forest blocks to a file geodatabase, a shapefile, and/or a KML file.
#
# Author: Daniel Otto
# ---------------------------------------------------------------------------

# Import modules
import arcpy, os, sys, zipfile

# Set overwrite output to true
arcpy.env.overwriteOutput = True

# Delete in-memory workspace
arcpy.Delete_management("in_memory")

# Get script path
Processing_Variables = {}
PYTHON_SCRIPT = sys.argv[0]
Processing_Variables['Script_Directory_Path'] = os.path.split(PYTHON_SCRIPT)[0]

# Get script arguments
name = sys.argv[1]
outPath = sys.argv[2]
newFolder = name + "_FNCS_upload"
Output_File_GDB = name + ".gdb"

# Set the qualified field names to false to match field names in both databases
arcpy.env.qualifiedFieldNames = False

# Define subset expression based on select_type and identifier
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

# Print expression
arcpy.AddMessage(Expression)

# Get desired output types
exportgdb = sys.argv[5]
exportshp = sys.argv[6]
exportKML = sys.argv[7]

# Set workspace to in-memory
arcpy.env.workspace = "in_memory"

def Create_folder():
    # Create new folder for output data
    arcpy.AddMessage("Creating new output folder...")
    Processing_Variables['Output_folder'] = outPath + '\\' + newFolder
    if arcpy.Exists(Processing_Variables['Output_folder']):
        msg = "There is already a workspace called " + newFolder + "."
        arcpy.AddMessage("----------------------------------------")
        arcpy.AddMessage(msg)
        arcpy.AddMessage("----------------------------------------")
        sys.exit()
    else:
        arcpy.CreateFolder_management(outPath,newFolder)

def Initialize():
    Processing_Variables['Blocks'] = "Database Connections\DBP06.sde\FORESTVIEW.SV_BLOCK"

def copy_features():
    # Select and merge the blocks into a single feature
    arcpy.Select_analysis(Processing_Variables['Blocks'],r'in_memory/blocks',Expression)

    # Get the count of selected features
    num_features = int(arcpy.GetCount_management(r'in_memory/blocks').getOutput(0))

    # Add a message indicating how many features were selected
    arcpy.AddMessage(str(num_features) +" Blocks selected.")

    # Delete all the columns besides the mandatory ones
    mandatory_fields = ["OBJECTID", "SHAPE"]
    delete_fields = [field.name for field in arcpy.ListFields(r'in_memory/blocks') if not field.required and field.type not in ["OID", "Geometry"]]
    delete_fields = list(set(delete_fields) - set(mandatory_fields))
    if len(delete_fields) > 0:
        arcpy.DeleteField_management(r'in_memory/blocks', delete_fields)

    # Dissolve all polygons into a single polygon
    arcpy.Dissolve_management(r'in_memory/blocks', r'in_memory/dissolved_blocks')

    # Select the dissolved polygon and delete the original blocks feature
    arcpy.Select_analysis(r'in_memory/dissolved_blocks', r'in_memory/combined_blocks')
    arcpy.Delete_management(r'in_memory/blocks')

def copy_2_FC():
    # Create output FGDB
    arcpy.CreateFileGDB_management(Processing_Variables['Output_folder'],"FNCS_upload.gdb","Current")
    Processing_Variables['Output_gdb'] = Processing_Variables['Output_folder'] +"\\FNCS_upload.gdb"

    # Copy features to output FGDB
    arcpy.CopyFeatures_management(r'in_memory/combined_blocks',Processing_Variables['Output_gdb'] + r'\Block')

def copy_2_SHP():
    # Create output shapefile folder
    arcpy.CreateFolder_management(Processing_Variables['Output_folder'],"Shapefiles")
    Processing_Variables['Shapefile_folder'] = Processing_Variables['Output_folder'] + '\\Shapefiles'

    # Copy features to output shapefile
    arcpy.CopyFeatures_management(r'in_memory/combined_blocks',Processing_Variables['Shapefile_folder'] + r'\Block.shp')

def kmz_to_kml(kmz_file_path, output_folder_path):
    with zipfile.ZipFile(kmz_file_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.endswith('.kml'):
                kml_file_name = os.path.splitext(os.path.basename(kmz_file_path))[0] + '.kml'
                kml_file_path = os.path.join(output_folder_path, kml_file_name)
                zip_ref.extract(file_name, output_folder_path)
                os.rename(os.path.join(output_folder_path, file_name), kml_file_path)
                return kml_file_path

def copy_2_KML():
    # Add layer to MXD, apply desired symbology, and export map to KMZ file
    input_fc = arcpy.MakeFeatureLayer_management(r'in_memory/combined_blocks', name)

    # Define the output KMZ file
    output_kmz = Processing_Variables['Output_folder'] +"\\"+ name +".kmz"

    # Convert layer to KML format and save as KMZ file
    arcpy.LayerToKML_conversion(input_fc, output_kmz)

    # Convert KMZ file to KML file
    kml_file_path = Processing_Variables['Output_folder']
    kmz_to_kml(output_kmz, kml_file_path)

    # Remove KMZ file
    os.remove(output_kmz)

if __name__ == '__main__':
    Create_folder()
    Initialize()
    copy_features()
    copy_2_FC()

    # Check if shapefile export is desired
    if exportshp == 'true':
        copy_2_SHP()

    # Check if KML export is desired
    if exportKML == 'true':
        copy_2_KML()

    # Check if GDB export is not desired
    if exportgdb == 'false':
        arcpy.Delete_management(Processing_Variables['Output_gdb'])

