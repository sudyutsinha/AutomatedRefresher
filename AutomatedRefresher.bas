Attribute VB_Name = "Module1"
Sub AutomatedRefresher()
' This Macro automates the process of refreshing FactSet FDS codes in multiple workbooks

    ' Below two lines declare variables to store the paths of the specific folder as well as the workbooks in the folder
    Dim fileName As String
    Dim folderName As String
    
    ' Below line stores the specific path in variable folderName, can be changed as desired
    folderName = ("C:\Users\sudyu\OneDrive\Documents\Macro Test\")
    
    ' Below line utilizes the directory function to get the excel workbooks in the path stored in folderName variable
    fileName = Dir(folderName & "*.xls*")
    
    ' Below line sets up a conditional loop that runs for every excel workbook in the file path
    Do While fileName <> ""
    
        ' Below line opens the excel workbook
        With Workbooks.Open(folderName & fileName)
            
            ' Below line refreshes all the FactSet FDS codes in the workbook
            ExecuteExcel4Macro "FDSFORCERECALC(FALSE)"
            ' Below line saves and closes the workbook
            ActiveWorkbook.Close SaveChanges:=True
            
        End With
    
    ' Below line is used to access the subsequent workbook in the folder path
    fileName = Dir()
    Loop

End Sub
