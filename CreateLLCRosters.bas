Attribute VB_Name = "Module1"
Sub CreateLLCRosters()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim savePath As String
    Dim fileNames As Variant
    Dim i As Integer
    Dim folderPath As String
    Dim shellApp As Object
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim dataRange As Range

    ' Set the path where the files will be saved
    Set shellApp = CreateObject("WScript.Shell")
    folderPath = shellApp.SpecialFolders("Desktop") & "\LLC Rosters\"
    
    ' Create the folder if it does not exist
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Set the active sheet to a variable
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Step 1: Delete specific columns (except for column C)
    ws.Columns("O:O").Delete
    ws.Columns("N:N").Delete
    ws.Columns("M:M").Delete
    ws.Columns("L:L").Delete
    ws.Columns("K:K").Delete
    ws.Columns("I:I").Delete
    ws.Columns("E:E").Delete
    ws.Columns("D:D").Delete
    
    ' Step 2: Create the list of file names
    fileNames = Array( _
        "LLC_Ally_Roster.xlsx", "LLC_Ally_Break_Roster.xlsx", "LLC_Arts_Architecture_Roster.xlsx", _
        "LLC_BIOME_Roster.xlsx", "LLC_BASH_Roster.xlsx", "LLC_EMS_Roster.xlsx", "LLC_EARTH_Roster.xlsx", _
        "LLC_ED_EQUITY_Roster.xlsx", "LLC_EHOUSE_Roster.xlsx", "LLC_FY_Education_Roster.xlsx", _
        "LLC_FY_Liberal_Arts_Roster.xlsx", "LLC_FISE_Roster.xlsx", "LLC_FY_Veterans_Roster.xlsx", _
        "LLC_Forensics_Roster.xlsx", "LLC_Flourish_Roster.xlsx", "LLC_Global_Engagement_Roster.xlsx", _
        "LLC_IST_House_Roster.xlsx", "LLC_ROTC_Roster.xlsx", "LLC_Paterno_Fellows_Roster.xlsx", _
        "LLC_PGM_Roster.xlsx", "LLC_Schreyer_Honors_Housing_Roster.xlsx", "LLC_GLOBE_Roster.xlsx", _
        "LLC_WISE_Roster.xlsx", "LLC_ROAR_Roster.xlsx", "LLC_Millennium_Scholars_Roster.xlsx" _
    )
    
    ' Step 3: Create and save the Excel Files
    For i = LBound(fileNames) To UBound(fileNames)
        ' Copy all data from the main worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Get last row of data
        Set dataRange = ws.Range("A1:Z" & lastRow) ' Adjust the range as needed (e.g., A to Z)
        
        ' Create a new workbook and paste the data
        dataRange.Copy
        Set newWorkbook = Workbooks.Add
        Set newWorksheet = newWorkbook.Sheets(1)
        newWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues ' Paste only values
        



' Apply filter for "LLC_Ally_Roster.xlsx" file
If fileNames(i) = "LLC_Ally_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC ALLY*", Operator:=xlOr, Criteria2:="=*UC LLC ALLY*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteAlly As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteAlly Is Nothing Then
                Set rowsToDeleteAlly = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteAlly = Union(rowsToDeleteAlly, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteAlly Is Nothing Then
        rowsToDeleteAlly.Delete
    End If
End If

' Apply filter for "LLC_Ally_Break_Roster.xlsx" file
If fileNames(i) = "LLC_Ally_Break_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally Break
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC ALLY BREAK*", Operator:=xlOr, Criteria2:="=*UC LLC ALLY BREAK*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteAllyBreak As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally Break
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteAllyBreak Is Nothing Then
                Set rowsToDeleteAllyBreak = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteAllyBreak = Union(rowsToDeleteAllyBreak, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteAllyBreak Is Nothing Then
        rowsToDeleteAllyBreak.Delete
    End If
End If

' Apply filter for "LLC_Arts_Architecture_Roster.xlsx" file
If fileNames(i) = "LLC_Arts_Architecture_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Arts & Architecture
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC A&A*", Operator:=xlOr, Criteria2:="=*UC LLC A&A*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteAA As Range
    
    ' Loop through the rows and add the hidden ones to the delete range
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteAA Is Nothing Then
                Set rowsToDeleteAA = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteAA = Union(rowsToDeleteAA, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteAA Is Nothing Then
        rowsToDeleteAA.Delete
    End If
End If

' Apply filter for "LLC_BIOME_Roster.xlsx" file
If fileNames(i) = "LLC_BIOME_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC BIOME*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteBiome As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Biome
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteBiome Is Nothing Then
                Set rowsToDeleteBiome = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteBiome = Union(rowsToDeleteBiome, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteBiome Is Nothing Then
        rowsToDeleteBiome.Delete
    End If
End If

' Apply filter for "LLC_BASH_Roster.xlsx" file
If fileNames(i) = "LLC_BASH_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC BASH*", Operator:=xlOr, Criteria2:="=*UC LLC BASH*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteBASH As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for BASH
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteBASH Is Nothing Then
                Set rowsToDeleteBASH = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteBASH = Union(rowsToDeleteBASH, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteBASH Is Nothing Then
        rowsToDeleteBASH.Delete
    End If
End If
       


' Apply filter for "LLC_EMS_Roster.xlsx" file
If fileNames(i) = "LLC_EMS_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for EMS
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC EMS*", Operator:=xlOr, Criteria2:="=*UC LLC EMS HOUSE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteEMS As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for EMS
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteEMS Is Nothing Then
                Set rowsToDeleteEMS = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteEMS = Union(rowsToDeleteEMS, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteEMS Is Nothing Then
        rowsToDeleteEMS.Delete
    End If
End If

' Apply filter for "LLC_EARTH_Roster.xlsx" file
If fileNames(i) = "LLC_EARTH_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for EARTH
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC EARTH*", Operator:=xlOr, Criteria2:="=*UC LLC EARTH*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteEARTH As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally Break
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteEARTH Is Nothing Then
                Set rowsToDeleteEARTH = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteEARTH = Union(rowsToDeleteEARTH, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteEARTH Is Nothing Then
        rowsToDeleteEARTH.Delete
    End If
End If

' Apply filter for "LLC_ED_EQUITY_Roster.xlsx" file
If fileNames(i) = "LLC_ED_EQUITY_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Arts & Architecture
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC ED EQUITY*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteED As Range
    
    ' Loop through the rows and add the hidden ones to the delete range
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteED Is Nothing Then
                Set rowsToDeleteED = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteED = Union(rowsToDeleteED, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteED Is Nothing Then
        rowsToDeleteED.Delete
    End If
End If

' Apply filter for "LLC_EHOUSE_Roster.xlsx" file
If fileNames(i) = "LLC_EHOUSE_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for EHOUSE
     newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC E-HOUSE*", Operator:=xlOr, Criteria2:="=*UC LLC E-HOUSE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteEHOUSE As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for EHOUSE
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteEHOUSE Is Nothing Then
                Set rowsToDeleteEHOUSE = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteEHOUSE = Union(rowsToDeleteEHOUSE, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteEHOUSE Is Nothing Then
        rowsToDeleteEHOUSE.Delete
    End If
End If

' Apply filter for "LLC_FY_Education_Roster.xlsx" file
If fileNames(i) = "LLC_FY_Education_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC EDUCATION*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteEDU As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for BASH
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteEDU Is Nothing Then
                Set rowsToDeleteEDU = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteEDU = Union(rowsToDeleteEDU, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteEDU Is Nothing Then
        rowsToDeleteEDU.Delete
    End If
End If



' Apply filter for "LLC_FY_Liberal_Arts_Roster.xlsx" file
If fileNames(i) = "LLC_FY_Liberal_Arts_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC Liberal Arts*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteLA As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteLA Is Nothing Then
                Set rowsToDeleteLA = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteLA = Union(rowsToDeleteLA, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteLA Is Nothing Then
        rowsToDeleteLA.Delete
    End If
End If

' Apply filter for "LLC_FISE_Roster.xlsx" file
If fileNames(i) = "LLC_FISE_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally Break
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC FISE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteFISE As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally Break
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteFISE Is Nothing Then
                Set rowsToDeleteFISE = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteFISE = Union(rowsToDeleteFISE, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteFISE Is Nothing Then
        rowsToDeleteFISE.Delete
    End If
End If

' Apply filter for "LLC_FY_Veterans_Roster.xlsx" file
If fileNames(i) = "LLC_FY_Veterans_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Arts & Architecture
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC VETERAN*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteVET As Range
    
    ' Loop through the rows and add the hidden ones to the delete range
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteVET Is Nothing Then
                Set rowsToDeleteVET = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteVET = Union(rowsToDeleteVET, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteVET Is Nothing Then
        rowsToDeleteVET.Delete
    End If
End If

' Apply filter for "LLC_Forensics_Roster.xlsx" file
If fileNames(i) = "LLC_Forensics_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC FORENSICS*", Operator:=xlOr, Criteria2:="=*UC LLC FORENSICS*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteFORENSICS As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Biome
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteFORENSICS Is Nothing Then
                Set rowsToDeleteFORENSICS = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteFORENSICS = Union(rowsToDeleteFORENSICS, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteFORENSICS Is Nothing Then
        rowsToDeleteFORENSICS.Delete
    End If
End If

' Apply filter for "LLC_Flourish_Roster.xlsx" file
If fileNames(i) = "LLC_Flourish_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC FLOURISH*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteFLOURISH As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for BASH
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteFLOURISH Is Nothing Then
                Set rowsToDeleteFLOURISH = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteFLOURISH = Union(rowsToDeleteFLOURISH, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteFLOURISH Is Nothing Then
        rowsToDeleteFLOURISH.Delete
    End If
End If




' Apply filter for "LLC_Global_Engagement_Roster.xlsx" file
If fileNames(i) = "LLC_Global_Engagement_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC GLOBAL ENGAGEMENT*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteGLO As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteGLO Is Nothing Then
                Set rowsToDeleteGLO = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteGLO = Union(rowsToDeleteGLO, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteGLO Is Nothing Then
        rowsToDeleteGLO.Delete
    End If
End If

' Apply filter for "LLC_IST_House_Roster.xlsx" file
If fileNames(i) = "LLC_IST_House_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally Break
     newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC IST*", Operator:=xlOr, Criteria2:="=*UC LLC IST HOUSE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteIST As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally Break
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteIST Is Nothing Then
                Set rowsToDeleteIST = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteIST = Union(rowsToDeleteIST, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteIST Is Nothing Then
        rowsToDeleteIST.Delete
    End If
End If

' Apply filter for "LLC_ROTC_Roster.xlsx" file
If fileNames(i) = "LLC_ROTC_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Arts & Architecture
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC TRI-SERVICE ROTC*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteROTC As Range
    
    ' Loop through the rows and add the hidden ones to the delete range
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteROTC Is Nothing Then
                Set rowsToDeleteROTC = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteROTC = Union(rowsToDeleteROTC, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteROTC Is Nothing Then
        rowsToDeleteROTC.Delete
    End If
End If

' Apply filter for "LLC_Paterno_Fellows_Roster.xlsx" file
If fileNames(i) = "LLC_Paterno_Fellows_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC PATERNO FELLOWS*"
    
    ' Store rows to be deleted
    Dim rowsToDeletePAT As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Biome
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeletePAT Is Nothing Then
                Set rowsToDeletePAT = newWorksheet.Rows(Row)
            Else
                Set rowsToDeletePAT = Union(rowsToDeletePAT, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeletePAT Is Nothing Then
        rowsToDeletePAT.Delete
    End If
End If

' Apply filter for "LLC_PGM_Roster.xlsx" file
If fileNames(i) = "LLC_PGM_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC PGM*", Operator:=xlOr, Criteria2:="=*UC LLC PGM*"
    
    ' Store rows to be deleted
    Dim rowsToDeletePGM As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for BASH
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeletePGM Is Nothing Then
                Set rowsToDeletePGM = newWorksheet.Rows(Row)
            Else
                Set rowsToDeletePGM = Union(rowsToDeletePGM, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeletePGM Is Nothing Then
        rowsToDeletePGM.Delete
    End If
End If





' Apply filter for "LLC_Schreyer_Honors_Housing_Roster.xlsx" file
If fileNames(i) = "LLC_Schreyer_Honors_Housing_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC SCHREYER*", Operator:=xlOr, Criteria2:="=*UC LLC SCHREYER*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteSCH As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteSCH Is Nothing Then
                Set rowsToDeleteSCH = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteSCH = Union(rowsToDeleteSCH, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteSCH Is Nothing Then
        rowsToDeleteSCH.Delete
    End If
End If

' Apply filter for "LLC_GLOBE_Roster.xlsx" file
If fileNames(i) = "LLC_GLOBE_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Ally Break
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC SCHREYER GLOBE*", Operator:=xlOr, Criteria2:="=*UC LLC SCHREYER GLOBE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteGLOBE As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Ally Break
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteGLOBE Is Nothing Then
                Set rowsToDeleteGLOBE = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteGLOBE = Union(rowsToDeleteGLOBE, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteGLOBE Is Nothing Then
        rowsToDeleteGLOBE.Delete
    End If
End If

' Apply filter for "LLC_WISE_Roster.xlsx" file
If fileNames(i) = "LLC_WISE_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Arts & Architecture
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY LLC WISE*", Operator:=xlOr, Criteria2:="=*UC LLC WISE*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteWISE As Range
    
    ' Loop through the rows and add the hidden ones to the delete range
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteWISE Is Nothing Then
                Set rowsToDeleteWISE = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteWISE = Union(rowsToDeleteWISE, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteWISE Is Nothing Then
        rowsToDeleteWISE.Delete
    End If
End If

' Apply filter for "LLC_ROAR_Roster.xlsx" file
If fileNames(i) = "LLC_ROAR_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*UC APT WHITE COURSE ROAR*"
    
    ' Store rows to be deleted
    Dim rowsToDeleteROAR As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for Biome
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteROAR Is Nothing Then
                Set rowsToDeleteROAR = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteROAR = Union(rowsToDeleteROAR, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteROAR Is Nothing Then
        rowsToDeleteROAR.Delete
    End If
End If

' Apply filter for "LLC_Millennium_Scholars_Roster.xlsx" file
If fileNames(i) = "LLC_Millennium_Scholars_Roster.xlsx" Then
    ' Apply the filter to column D with the specified criteria for Biome
    newWorksheet.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:="=*FY MILLENNIUM*", Operator:=xlOr, Criteria2:="=*UC MILLENNIUM (SIMMONS)*"
     
    ' Store rows to be deleted
    Dim rowsToDeleteMILL As Range
    
    ' Loop through the rows and add the hidden ones to the delete range for BASH
    For Row = 2 To lastRow ' Assuming row 1 is the header
        If newWorksheet.Rows(Row).Hidden = True Then
            If rowsToDeleteMILL Is Nothing Then
                Set rowsToDeleteMILL = newWorksheet.Rows(Row)
            Else
                Set rowsToDeleteMILL = Union(rowsToDeleteMILL, newWorksheet.Rows(Row))
            End If
        End If
    Next Row
    
    ' Delete all the rows at once if there are any to delete
    If Not rowsToDeleteMILL Is Nothing Then
        rowsToDeleteMILL.Delete
    End If
End If

        ' Save the new file
        newWorkbook.SaveAs folderPath & fileNames(i)
        newWorkbook.Close False
    Next i
    
    MsgBox "All files have been saved successfully in the 'LLC Rosters' folder on your desktop.", vbInformation
End Sub

