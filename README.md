FindDuplicatesAddIn
===================

Find Duplicates tool for ESRI ArcMap use, written in VB.Net. 
Supplied as a Visual Studio 2010 Solution.

ESRI Add-In does not require admin rights to install.
Requires Visual Studio and ArcMap SDK to build.

This code is based off of VBA code downloaded from the ESRI site back in 2010, unfortunately I do not have the attribution information of the original writer and am now unable to find the VBA code on the ESRI site. If you know who wrote this please let me know so that I can attribute them correctly.

------------------
Installing AddIns:

Copy the .esriAddIn file to a local folder, or the well known folder below:
Vista/7: C:\Users\<username>\Documents\ArcGIS\AddIns\Desktop10.1
XP: C:\Documents and Settings\<username>\My Documents\ArcGIS\AddIns\Desktop10.1

Open the appropriate Arc product (eg ArcMap, ArcGIS Explorer). 
Choose Customize, Add-In Manager.
If you placed the Add-In in a well known location then it should appear in the My Add-Ins section. If it is elsewhere, click the options tab and add the folder location that you placed the AddIn into. Make sure that you have selected the “Load all Add-ins without restrictions” radio button as this add-in is not digitally signed.
Click Close and reopen the Add-In Manager.
You should now be able to see all the add-ins in the Shared Add-ins or My Add-ins section.

Click the Customize button and choose the commands tab.
In the Add-In Controls category you will find the various add-ins available, drag them to a toolbar to access them.
