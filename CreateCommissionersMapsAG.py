"""Create Commissioners Maps
        Extract Data
        Repath and Update Maps
        Export to PDF's
        Create LON Info
Version 2.1 - Update for use by Ag staff
Bryan Grill - Lancaster County IT/GIS
January 2016"""

#Updates:
#All data besides excluded/edited farm is now in temp data gdb
#******************************************************IMPORT MODULES, GET USER INFO**************************************************************************
import arcpy, os, sys, shutil, getpass, subprocess
from arcpy import env
from datetime import datetime

user = getpass.getuser()
now = datetime.now()

#Overwrite Workspace Data if necessary
arcpy.env.overwriteOutput = True

#Does user have Image module?
destination = r"local\path\to\python\PIL"
if os.path.exists(destination):
    print "Requested module exists"
else:
    subprocess.Popen(r"\PIL\module\on\network\PIL-1.1.7.win32-py2.7.exe")
    print "Install Image Module"
    sys.exit()

import Image
#******************************************************PULL DATA FARM, CLIP, DISSOLVE, ECT.********************************************************************

# Workspace Vars
root = r"project\root"
tempOutput = os.path.join(root, "CommTempdata")
output = r"\output\data\onto\network\here"
tempWorkspace = os.path.join(tempOutput, "CommissionersTemp.gdb")
tempMaps = os.path.join(tempOutput, "Maps")

#User Inputs Vars
farminput = arcpy.GetParameterAsText(0)
farminput = farminput.replace(',','')
filenum = arcpy.GetParameterAsText(1)
easeac = arcpy.GetParameterAsText(2)
ownername = arcpy.GetParameterAsText(3)

#Data Input Vars
SDE_parcel = r"\path\to\data\parcel_poly"
SDE_greenland = r"\path\to\data\greenland"
SDE_muni = r"\path\to\data\Muni"
soils_fullname = r"\path\to\data\Soils_2013"

#Farm Variables - Accepts up to four accounts
if len (farminput) > 26:
    farm0,farm1,farm2,farm3 = farminput.split()
    farm0 = str(farm0[0:8])
    farm1 = str(farm1[0:8])
    farm2 = str(farm2[0:8])
    farm3 = str(farm3[0:8])
    farmQuery = '"ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'' % (farm0, farm1, farm2, farm3)
elif len(farminput) <= 26 and len(farminput) > 17:
    farm0,farm1,farm2 = farminput.split()
    farm0 = str(farm0[0:8])
    farm1 = str(farm1[0:8])
    farm2 = str(farm2[0:8])
    farmQuery = '"ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'' % (farm0, farm1, farm2)
elif len(farminput) <= 17 and len(farminput) > 8:
    farm0,farm1 = farminput.split()
    farm0 = str(farm0[0:8])
    farm1 = str(farm1[0:8])
    farmQuery = '"ACCOUNT" = \'%s00000\'OR "ACCOUNT" = \'%s00000\'' % (farm0, farm1)
else:
    farm0 = farminput
    farm0 = str(farm0[0:8])
    farmQuery = '"ACCOUNT" = \'%s00000\'' % (farm0)

arcpy.AddMessage(farminput)

#Temp Vars
tempsoil = os.path.join(tempWorkspace, "tempsoil")
tempgreenland = os.path.join(tempWorkspace, "tempgreenland")
tempgreenland_lu = os.path.join(tempWorkspace, "tempgreenland_lu")
farmSelect = os.path.join(tempWorkspace, "farmtemp1")
farmSelect1 = os.path.join(tempWorkspace, "farmtemp2")
farmExclusion = os.path.join(tempWorkspace, "FarmExclusion")
farmFinal = os.path.join(tempWorkspace, "Farm")
soilbwcolor = os.path.join(tempWorkspace, "soil%s") % (farm0)
soilcalc = os.path.join(tempWorkspace, "soilcalc%s") % (farm0)
soilLU = os.path.join(tempWorkspace, "soilLU%s") % (farm0)
ParcelLyrTemp = "ParcelLayerTemp"
LONtemp = os.path.join(tempWorkspace, "LON")

#Create or Delete/Create Workspace Folder
def workspaceFunc():
    if not os.path.exists(root): os.makedirs(root)
    try:
        if os.path.isfile(tempOutput):
            os.remove(tempOutput)
            arcpy.CreateFolder_management(root, "CommTempdata")
            arcpy.CreateFileGDB_management(tempOutput, "CommissionersTemp", "")
        else:
            arcpy.CreateFolder_management(root, "CommTempdata")
            arcpy.CreateFileGDB_management(tempOutput, "CommissionersTemp", "")
    except arcpy.ExecuteError:
        pass

#Function to select farm and confirm it exists
def selectFarmFunc(source, out, query):
    arcpy.Select_analysis(source, out, query)
    #Confirm account exists
    featureCount = arcpy.GetCount_management(out)
    if str(featureCount) < str(1):
        arcpy.Delete_management(out, "FeatureClass")
        print "Account Does Not Exist as entered"
        arcpy.AddMessage("Account Does Not Exist as entered")
        sys.exit()
    else:
        print "Farm selected and extracted"
        arcpy.AddMessage("Farm(s) selected and extracted")

#create the workspace and get the farm
workspaceFunc()
selectFarmFunc(SDE_parcel, farmSelect1, farmQuery)
farmSelect = farmSelect1

#Clip Soils to farm and dissolve for Soil Color/BW Maps
arcpy.Clip_analysis(soils_fullname, farmSelect, tempsoil)
arcpy.Dissolve_management(tempsoil, soilbwcolor, "MUSYM", "CLNIRR FIRST;FULLNAME FIRST", "MULTI_PART", "DISSOLVE_LINES")

#Clip Greenland to farm and dissolve on Unit field for SoilCalc Map - Soil Calculations table
arcpy.Clip_analysis(SDE_greenland, farmSelect, tempgreenland)
arcpy.Dissolve_management(tempgreenland, soilcalc, "UNIT", "", "MULTI_PART", "DISSOLVE_LINES")
arcpy.AddField_management(soilcalc, "ACRES", "FLOAT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.CalculateField_management(soilcalc, "ACRES", "!shape.area@acres!", "PYTHON_9.3", "")

#Clip Greenland to farm and dissolve on Soilland field for SoilCalc Map - Soil/Landuse table
arcpy.Copy_management(tempgreenland, tempgreenland_lu, "")
arcpy.AddField_management(tempgreenland_lu, "SOILLAND", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.CalculateField_management(tempgreenland_lu, "SOILLAND", "[CLASS] & [USE_]", "VB", "")
arcpy.Dissolve_management(tempgreenland_lu, soilLU, "SOILLAND", "", "MULTI_PART", "DISSOLVE_LINES")
arcpy.AddField_management(soilLU, "ACRES", "FLOAT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.CalculateField_management(soilLU, "ACRES", "!shape.area@acres!", "PYTHON_9.3", "")

#Merge all four features classes, move other three to temp directory
arcpy.Merge_management([farmSelect, soilbwcolor, soilcalc, soilLU], farmFinal)

#Drop Unnecessary fields
deletefields1 = "FIELDS;TO;DELETE"
arcpy.DeleteField_management(farmFinal, deletefields1)

#Create Folder by File number to house maps LON for farm
commissionersfolder = r"\path\to\output"
farmFolder= os.path.join(commissionersfolder, "%s") % (filenum)

if os.path.exists(farmFolder):
    pass
else:
    arcpy.CreateFolder_management(commissionersfolder, filenum)

#********************************************************************************UPDATE MAP TEMPLATES************************************************************************************
#Pull Owner Last Name
if str(ownername) == "":
    name = arcpy.SearchCursor(farmFinal, "", "", "OWNER_NAME", "OWNER_NAME")
    for row in name:
        nameAll = row.OWNER_NAME
    lastName, rest = nameAll.split(" ", 1)
    lastName = lastName.title()
else:
    lastName = ownername.title()


lastName = lastName.title()
lastNameLayout = "The " + lastName + " Farm"
lastNameLayout2 = lastName + " Farm"
legendname = lastName + " Parcel"

#Pull Muni District
munithreedigit = farm0[0:3]
core = farm0[3:8]

acctdisplay = munithreedigit + " - " + core + " - " + "0" + " - " + "0000"

if munithreedigit == "461":
    munithreedigit = "460"

muniquery = '"DISTRICT" = \'%s\'' % (munithreedigit)

#Pull Muni Name
muniName = arcpy.SearchCursor(SDE_muni, muniquery, "", "MUNICIPALITY", "MUNICIPALITY")
for row in muniName:
    muni = row.MUNICIPALITY
    muni = muni.title()

#Pull GIS acres
getacres = arcpy.SearchCursor(farmFinal, "", "", "DEED_AREA", "DEED_AREA")
for row in getacres:
    gisacres = row.DEED_AREA

#Repath data in templates
farmname = "farm%s" % (farm0)
LONtemp2 = "LON"

#CREATE LON
#output Variables
LONtable = "\\\\Nq-cluster2\\share-AG\\Commissioners' Maps\\%s\\LON_%s.xls" % (filenum, lastName)
fieldstokeep = "FIELDS;TO;KEEP"
arcpy.MakeFeatureLayer_management(SDE_parcel, ParcelLyrTemp, "", tempOutput, fieldstokeep)
arcpy.SelectLayerByLocation_management(ParcelLyrTemp, "WITHIN_A_DISTANCE", farmFinal, "125 feet", "NEW_SELECTION" )
arcpy.CopyFeatures_management(ParcelLyrTemp, LONtemp)
arcpy.TableToExcel_conversion(LONtemp, LONtable, "ALIAS")

#Map Production Function
def MakeMaps(comm, sketch, twp, truecolor, soilcolor, soilbw, soilcalc, lon):
    arcpy.AddMessage("Making Maps")
    #*****UPDATE LOCATION MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(comm)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    df2 = arcpy.mapping.ListDataFrames(mxd, "LocationMap")[0]
    layer1 = "Farm"
    layer1a = "Parcel"
    layer1b = "SelectedMunicipality"
    layer1c = "MinorRoads"
    pdf = farmFolder + "\\Location_%s.pdf" % (lastName)
    #Queries
    distzoom = '"DISTRICT" = \'%s\'' % (munithreedigit)
    roadpath = r"\path\to\data\by\muni\d%s" % (munithreedigit)
    #Update Layers
    for lyr2 in arcpy.mapping.ListLayers(mxd, layer1c, df):
        lyr2.replaceDataSource(roadpath, "SHAPEFILE_WORKSPACE", "rdcline")
    for lyr3 in arcpy.mapping.ListLayers(mxd, layer1a, df):
        lyr3.definitionQuery = distzoom
    for lyr4 in arcpy.mapping.ListLayers(mxd, layer1b, df):
        lyr4.definitionQuery = distzoom
    for lyr5 in arcpy.mapping.ListLayers(mxd, layer1b, df2):
        lyr5.definitionQuery = distzoom
    newextent = lyr5.getExtent()
    df.extent = newextent
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    updatetext2 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName1")[0]
    updatetext2.text = lastNameLayout
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName2")[0]
    updatetext3.text = lastNameLayout
    updatetext4 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "FILENO")[0]
    updatetext4.text = filenum
    updatetext5 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "MUNI")[0]
    updatetext5.text = muni
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName3")[0]
    updatetext3.text = lastNameLayout2
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE TAX MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(sketch)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    pdf = farmFolder + "\\Tax_%s.pdf" % (lastName)
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer1, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale *= 1.5
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    if str(ownername) != "":
        updatetext2 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "NAME")[0]
        updatetext2.text = ownername
    else:
        updatetext2 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "NAME")[0]
        updatetext2.text = nameAll.title()
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "FILENO")[0]
    updatetext3.text = filenum
    updatetext4 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "GISAC")[0]
    updatetext4.text = gisacres
    if str(easeac) != "":
        updatetext5 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "EASEAC")[0]
        updatetext5.text = easeac
    else:
        updatetext5 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "EASEAC")[0]
        updatetext5.text = gisacres
    updatetext6 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LEGENDNAME")[0]
    updatetext6.text = legendname
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE TOPO MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(twp)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    pdf = farmFolder + "\\Topo_%s.pdf" % (lastName)
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer1, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale = 17500
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    updatetext2 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName1")[0]
    updatetext2.text = lastNameLayout
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName2")[0]
    updatetext3.text = lastNameLayout
    updatetext4 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "FILENO")[0]
    updatetext4.text = filenum
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE TRUE COLOR MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(truecolor)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    pdf = farmFolder + "\\TrueColor_%s.pdf" % (lastName)
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer1, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale *= 1.1
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE SOIL COLOR MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(soilcolor)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    layer2 = "Soils Color"
    pdf = farmFolder + "\\SoilColor_%s.pdf" % (lastName)
    soilcolorlyr = arcpy.mapping.ListLayers(mxd, "Soils Color", df)[0]
    soilrpt = r"\path\to\some\data\SoilsColorRpt.rlf"
    soilrpt_tif = r"\path\to\some\data\SoilsColorRpt.tif"
    soilrpt_jpg = r"\path\to\some\data\SoilsColorRpt.jpg"
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer2, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale *= 1.1
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName1")[0]
    updatetext3.text = lastNameLayout2
    #Update Table
    arcpy.mapping.ExportReport(soilcolorlyr, soilrpt, soilrpt_tif, "ALL")
    im = Image.open(soilrpt_tif)
    im.save(soilrpt_jpg)
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE SOIL BW MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(soilbw)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    layer3 = "Soils BW"
    pdf = farmFolder + "\\SoilBW_%s.pdf" % (lastName)
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer3, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale *= 1.1
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    updatetext3 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LastName1")[0]
    updatetext3.text = lastNameLayout2
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE SOIL CALC MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(soilcalc)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    layer4 = "Soil Calculations"
    layer5 = "Soil/Landuse"
    pdf = farmFolder + "\\SoilCalc_%s.pdf" % (lastName)
    soilcalcrpt = r"\path\to\some\data\SoilCalc.rlf"
    soilcalcrpt_tif = r"\path\to\some\data\SoilCalcRpt.tif"
    soilcalcrpt_jpg = r"\path\to\some\data\SoilsCalcRpt.jpg"
    soilLUrpt = r"\path\to\some\data\SoilLandUse.rlf"
    soilLUrpt_tif = r"\path\to\some\data\SoilLURpt.tif"
    soilLUrpt_jpg = r"\path\to\some\data\SoilLURpt.jpg"
    #Update Layers
    for lyr in arcpy.mapping.ListLayers(mxd, layer1, df):
        newextent = lyr.getExtent()
    df.extent = newextent
    df.scale *= 1.1
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    #Update Tables
    soilcalclyr = arcpy.mapping.ListLayers(mxd, "Soil Calculations", df)[0]
    arcpy.mapping.ExportReport(soilcalclyr, soilcalcrpt, soilcalcrpt_tif, "ALL")
    im = Image.open(soilcalcrpt_tif)
    im.save(soilcalcrpt_jpg)
    soil_lu_lyr = arcpy.mapping.ListLayers(mxd, "Soil/Landuse", df)[0]
    arcpy.mapping.ExportReport(soil_lu_lyr, soilLUrpt, soilLUrpt_tif, "ALL")
    im = Image.open(soilLUrpt_tif)
    im.save(soilLUrpt_jpg)
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

    #*****UPDATE LON MAP*****
    #Variables
    mxd = arcpy.mapping.MapDocument(lon)
    df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
    layer6 = "Notified Parcels"
    pdf = farmFolder + "\\LONmap_%s.pdf" % (lastName)
    #Update Layers
    for lyr2 in arcpy.mapping.ListLayers(mxd, layer6, df):
        newextent = lyr2.getExtent()
    df.extent = newextent
    df.scale *= 1.1
    #Update Text
    updatetext1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
    updatetext1.text = acctdisplay
    #Export to PDF
    arcpy.mapping.ExportToPDF(mxd, pdf)
    mxd.save()

#Standard Template Pathes
comm1 = r"\path\to\some\mxds\1-Commissioner.mxd"
sketch2 = r"\path\to\some\mxds\2-Sketch_Tax.mxd"
twp3 = r"\path\to\some\mxds\3-Township_Topo.mxd"
truecolor4 = r"\path\to\some\mxds\4-TrueColor.mxd"
soilcolor5 = r"\path\to\some\mxds\5-Soils_color.mxd"
soilbw6 = r"\path\to\some\mxds\6-Soils_bw.mxd"
soilcalc7 = r"\path\to\some\mxds\7-SoilCalcMap.mxd"
lon8 = r"\path\to\some\mxds\8-LON.mxd"

#Copy maps to local so user can save
def copyMaps():
    mapDocList = [comm1, sketch2, twp3, truecolor4, soilcolor5, soilbw6, soilcalc7, lon8]
    os.makedirs(tempMaps)

    for mxd in mapDocList:
        shutil.copy(mxd, tempMaps)

copyMaps()

#Make the maps
commLocal = os.path.join(tempMaps, "1-Commissioner.mxd")
sketchLocal = os.path.join(tempMaps, "2-Sketch_Tax.mxd")
twpLocal = os.path.join(tempMaps, "3-Township_Topo.mxd")
truecolorLocal = os.path.join(tempMaps, "4-TrueColor.mxd")
soilcolorLocal = os.path.join(tempMaps, "5-Soils_color.mxd")
soilbwLocal = os.path.join(tempMaps, "6-Soils_bw.mxd")
soilcalcLocal = os.path.join(tempMaps, "7-SoilCalcMap.mxd")
lonLocal = os.path.join(tempMaps, "8-LON.mxd")

MakeMaps(commLocal, sketchLocal, twpLocal, truecolorLocal, soilcolorLocal, soilbwLocal, soilcalcLocal, lonLocal)

#Open Folder
arcpy.AddMessage(farmFolder)
os.startfile(farmFolder)


