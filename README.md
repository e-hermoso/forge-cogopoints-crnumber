# Rename Corner Record Forms, Layout, and COGO Points with record numbers

## Description

This repository contains C# project of an AutoCAD CRX plug-in for use with Design Automation for Civil 3D and Salesforce. The plug-in implements an AutoCAD custom command named OCPWRENAME. This command renames the layout name, corner record form, and COGO Points with the corner record number that is provided in a json format from Salesforce. The updated CAD drawing is then saved.

## Dependencies
-  [Visual Studio](https://visualstudio.microsoft.com/downloads/). This sample was created in Visual Studio 2017.
-  [AutoCAD .NET API](https://www.nuget.org/packages/AutoCAD.NET/23.1.0). This sample was built using the AutoCAD 2020 .NET API.

## CAD File and Json File
- The CAD File must include a template of a corner record form. Each Layout in the CAD drawing are named as CR1, CR2, CR3, ets., as well as COGO Points in the Name field. The Cad drawing must follow this specific standard and naming convention for the plug-in to work properly.
- The Json File contains an alias corner record name as the key and corner record number as the value ({“cr1”: “123-123”}). The alias corner record name in the json file is matched with the corner record name in the CAD drawing and then updated with the corner record number. The json file is generated from Orange County Salesforce API after a CAD drawing has been submitted and analyzed through the Orange County portal website using Forge Design Automation.

## Build the project

1. Open *LMSCornerRecordRename.sln* in Visual Studio.

2. In Visual Studio, in the Solution Explorer, right-click the project name. A menu displays.

3. Click **Manage NuGet Packages**. The Manage Packages for Solution dialog displays.

4. Click **Restore**, which is on the top-right of the dialog. The packages are downloaded and restored.

5. Build the solution. 
## Modify Code (optional)

1. Change the path that reads the json file.

## Test the custom command

We recommend that you test the custom command on your local machine before you use it in Design Automation.

1. Start AutoCAD. (This code sample was tested with AutoCAD 2020)

2. Open the drawing file that is provided in this SampleFiles folder.

3. Make sure the params.json file exist in the bin/Debug directory or modify the code that read the json file saved in the SampleFiles folder.

4. On the AutoCAD command line, enter NETLOAD.

5. Select *LMSCornerRecordRename.dll* that you built in the previous section.

6. On the AutoCAD command line, enter OCPWRENAME. If the plug-in executes as designed, the CAD drawing file will be updated and generated. The file is typically saved in the folder that you most recently interacted with in AutoCAD. Note before running the command make sure the json file is path correctly. 
