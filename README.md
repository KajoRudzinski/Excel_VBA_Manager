# Excel_VBA_Manager
The aim of the project is to provide code that will allow you to export/import VBA code (classes, modules, forms, etc.) into/from any chosen Excel file.

## Current status
I have uploaded a working module that let's you export all VBA components (class, module, form, sheet, worbook) from active workbook.

## How to use

#### Exporting code form active workbook
1. Import the module ExportCodeFromThisWorkbook.bas into VB Project containg code you wish to export. 
2. Make sure you are using VBA references listed below.
3. Make sure your VBA code is not protected.
4. Execute procedure ```ExportCodeFromThisWorkbook()```. You will be asked to choose a folder to export to and then informed upon the export result.

## VBA references used
- Visual Basic fo Applications
- Microsoft Excel 16.0 Object Library
- OLE Automation
- Microsoft Office 16.0 Object Library
- Microsoft Scripting Runtime
- Microsoft Visual Basic for Applications Extensibiliy 5.3

## Environment
- Office 365 personal 
- Windows 10

## Inspired by
- https://www.rondebruin.nl/
- https://www.youtube.com/user/WiseOwlTutorials
