## Excel-Add-in

A Sample app that will read in a specifically formatted Excel file and create a Part BOM structure inside of Innovator.

## Important!

Always back up your code tree and database before applying an import package or code tree patch!



## Pre-requisites

1. Aras Innovator installed (version 12.0 SPx preferred)
2. Aras Package Import tool
3. Visual Studio
	>Note you need to install toolsets: Web& Cloud -> Office/SharePoint development
4. Microsoft Office Excel
5. .NET\IOM.DLL

This DLL is used when developing .NET applications or modules.
e.g. The Package Import Export Utilities or a custom website that connects to the Innovator Server.



## Install Steps

1. Back up your database and store the BAK file in a safe place.
2. Open the Aras Package Import tool.
3. Enter your login credentials and click Login.
	>Note: You must log in as root for the package import to succeed!
4. Enter the package name in the TargetRelease field.
	>Optional: Enter a description in the Description field.
5. Enter the path to your local ..\Excel-Add-In\Aml\imports.mf file in the Manifest File field.
6. Select ExcelAddIn in the Available for Import field.
7. Select Type = Merge and Mode = Thorough Mode.
8. Click Import in the top-left corner.
9. Close the Aras Package Import tool.
10. Open the ExcelAddIn.sln solution by path: ..\Excel-Add-In\ExcelAddIn\ExcelAddIn.sln.
11. Add reference to IOM.DLL
12. Open file named ThisAddIn.cs.
13. Change the next fields according to your installed environment: userName, password, database, innovatorServerUrl.
14. Build solution.
15. Run file by path ..\Excel-Add-In\ExcelAddIn\ExcelAddIn\bin\Debug\ExcelAddIn.vsto.

## How It Works

![Excel-Add-In] (screenshots/excel-add-in.png.png)
![Part] (screenshots/part.png)

## Usage

1. Open ..\ExcelAddIn\test.xlsx file for view.
2. Go to "Custom Aras Labs Tab" tab.
	>Note: If the "Custom Aras Labs Tab" tab is not present than go to File -> Options -> Add-ins, select 'Disabled Items' for Manage and click 'Go...' button, select "Custom Aras Labs Tab" and enable it.
3. Click Import BOM Example 1 for sheet named Example1 and Import BOM Example 2 for sheet named Example2.
4. Log in to Aras as admin.
5. Navigate to Design/Parts in the table of contents (TOC).
6. Check if the items from test.xlsx present and check Bom structure according to excel data structure. 

## Credits

# Created by the Aras Labs team:
* Rob McAveney
* Christopher Gillis
* Volodymyr Shyshkivskyi
