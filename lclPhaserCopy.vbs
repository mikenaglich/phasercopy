' PhaserCopy.vbs  
'
' VBScript that can be called from file explorer or windows scheduler to execute the Excel macro 
' that distributes PGx PDF files stored in directory 'S:\MED\PHASER\PHASER Sites\@PGX Results From Sanford'
' to individual PHASER site directories.  The parameter worksheet in PhaserCopyV*.xlsm
' defines individual PHASER sites and where their PGx results will be copied.  
' 
' 2020-01-15 mdn Start on it.
' 2020-01-24 mdn First version complete.
' 2020-03-13 mdn Prompt user before executing the file copy.
' 2020-04-05 mdn Display log worksheet when PhaserCopyV2.xlsm finishes. 
' 2020-04-11 mdn Add GUI in PhaserCopyV3.xlsm. 
' 2020-04-13 mdn Add CDW report option in PhaserCopyV4.xlsm.
' 2020-04-26 mdn Point to V5.
' 2020-06-08 mdn Point to V6.
' 2020-12-08 mdn Point to V7.
' 2020-12-31 mdn Point to V8.
' 2021-01-02 mdn Point to V8a.
' 2021-01-17 mdn Point to V8b.
' 2021-05-03 mdn Point to V8b on local C: drive to improve performance.
' 2022-01-16 mdn Home version
'********************************************************************************
    'Input Excel File Full Path
    '    ExcelFilePath = "S:\MED\PHASER\PHASER Sites\@automation\PhaserCopyV8b.xlsm"   ' On va06   
        ExcelFilePath = "C:\Users\miken\MyData-PC\PCyV9\PhaserCopyV9.xlsm"    
    'Macro name within the Excel File
        MacroPath = "PGxCopy.PGxMain"
    'Create an instance of Excel
        Set ExcelApp = CreateObject("Excel.Application")
    'Turn off Excel message diplays
        ExcelApp.Visible = False  
	ExcelApp.ScreenUpdating = False 
        ExcelApp.DisplayAlerts = False
    'Open Excel File and run the macro code.
        Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)
        ExcelApp.Run MacroPath
    'Close Excel File
        ExcelApp.Visible = True 
        ExcelApp.ScreenUpdating = True 
        ExcelApp.Worksheets("LOG").Activate
    ' Leave with Excel still opened to the LOG worksheet.    
 
WScript.Quit