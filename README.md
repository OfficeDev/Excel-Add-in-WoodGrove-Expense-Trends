# Excel-Add-in-JS-WoodGrove-Expense-Trends

The WoodGrove Bank Expense Trends add-in demonstrates how you can use the new JavaScript API for Microsoft Excel 2016 to create a compelling Excel add-in. With Expense Trends, you can import expense transactions into the workbook, create dashboard and trackers, view and analyze trends, and track special transactions such as charitable donations and follow up items. The sample provides two experiences: one with task pane and another with add-in commands. The following figures show the main screens of this add-in.

![WoodGrove Bank Expense Trends Add-in - Ribbon] (images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank Expense Trends Add-in - Initial taskpane] (images/woodgrove_taskpane_import.PNG)

![WoodGrove Bank Expense Trends Add-in - Transactions sheet] (images/woodgrove_taskpane_data.PNG)

![WoodGrove Bank Expense Trends Add-in - Dashboard] (images/woodgrove_taskpane_dashboard.PNG)

![WoodGrove Bank Expense Trends Add-in - Donations Tracker] (images/woodgrove_taskpane_donations.PNG)

## Table of Contents

* [Prerequisites](#prerequisites)
* [Run the project](#run-the-project)
* [Additional resources](#additional-resources)

## Prerequisites

You'll need:

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools for Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* Excel 2016, version 6769.2011 or later

## Run the project

1. Copy the project to a local folder. Ensure that the file path is not too long, otherwise you might run into an error in Visual Studio when it tries to install the NuGet packages necessary for the project. 
2. Then open the `WoodGrove Expense Trends.sln` in Visual Studio. 
3. Press F5 to build and deploy the sample add-in. Excel launches and depending on the version of Excel 2016 you've, the add-in loads a custom tab called WoodGrove in the ribbon, or opens in a task pane to the right of the worksheet, as shown in the following figures.

![WoodGrove Bank Expense Trends Add-in - Initial taskpane] (images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank Expense Trends Add-in - Initial taskpane] (images/woodgrove_taskpane_import.PNG)

## Additional resources

* [Office Dev Center](http://dev.office.com/)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.

