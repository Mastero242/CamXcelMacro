Attribute VB_Name = "ImagesController"

'namespace=vba-files\Helpers

Option Explicit

Public RootPath As String

'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public  Sub SetRootPath()
  RootPath = GetFolder() + "\"
  MsgBox RootPath,,"" 
End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public  Sub CreateCatalogSheet()
  Dim sheet As Worksheet
  Dim sheetname As Variant

  ' detect if catalog sheet exist, if yes delete and create if no only create
  sheetname = "Catalog"
  Application.DisplayAlerts = False

  For Each sheet In ActiveWorkbook.Worksheets
    If sheetname = sheet.Name Then
      sheet.Delete
    End If
  Next sheet
  Application.DisplayAlerts = True

  Set sheet = ActiveWorkbook.Sheets.Add(Before:=Sheets(1))
  With sheet
    .Name = "Catalog"
  End With

  'import image
  Call ImportImages
End Sub

Public Sub ImportImages()

    Const factor = 0.9  'picture is 90% of the size of cell

    'Variable Declaration
    Dim fsoLibrary As FileSystemObject
    Dim fsoFolder As Object
    Dim sFolderPath As String
    Dim sFileName As Object
    Dim p As Object
    Dim pic As Picture

    Dim i As Long   'counter
    Dim last_row As Long
    Dim ws As Worksheet

    Dim colFolders As New Collection
    Dim oFolder As Object
    Dim sf As Object
      
    'sFolderPath = RootPath 'ActiveWorkbook.Path + "\JOOR\PERMANENTS\"  'may need to change this line to suit your situation

    If Trim(RootPath) <> "" Then
      sFolderPath = RootPath
    Else
      MsgBox "Root path is empty"
      End
    End If
    
    'Set all the references to the FSO Library
    Set fsoLibrary = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fsoLibrary.GetFolder(sFolderPath)
    Set ws = Application.ActiveSheet
    On Error Resume Next
    
    With ws
        .Range("A1") = "Root Folder"
        .Range("B1") = "Current Folder"
        .Range("C1") = "Name"
        .Range("D1") = "Path"
        '.Range("E1") = "Picture"

        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 15        
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 50

        i = 2
        colFolders.Add fsoFolder
        Do While colFolders.Count > 0
          Set oFolder = colFolders(1)
          colFolders.Remove 1

          .Cells(i, 1) = fsoFolder.Name
          .Cells(i, 2) = oFolder.Name

          Dim rangefoldername As Object
          Set rangefoldername = Range(Cells(i, 1), Cells(i, 5))

          With rangefoldername
            .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            .Interior.Color = RGB(200, 200, 200)
          End With

          i = i + 1

          'Loop through each file in a folder
          For Each sFileName In oFolder.Files
            
            .Cells(i, 3) = Left(sFileName.Name, InStr(sFileName.Name, ".") - 1)
            .Cells(i, 4) = sFileName

            .Cells(i, 3).WrapText = True
            .Cells(i, 4).WrapText = True
            
            ' Set p = .Shapes.AddPicture(Filename:=sFileName, LinkToFile:=False, SaveWithDocument:=True, _
            '     Left:=.Cells(i, 5).Left, Top:=Cells(i, 5).Left, Width:=-1, Height:=-1)

            ' p.Width = .Cells(i, 5).Width * factor
            
            ' 'adjust row height
            ' If .Cells(i, 5).RowHeight < p.Height / factor Then
            '     .Cells(i, 5).RowHeight = p.Height / factor
            ' End If

            ' p.Left = .Cells(i, 5).Left + (.Cells(i, 5).Width - p.Width) / 2
            ' p.Top = .Cells(i, 5).Top + (.Cells(i, 5).Height - p.Height) / 2

            i = i + 1

          Next sFileName

          'add any subfolders to the collection for processing
          For Each sf In oFolder.subfolders
            colFolders.Add sf 'add to collection for processing
          Next sf
        Loop
            
    End With
    
    'Release the memory
    Set fsoLibrary = Nothing
    Set fsoFolder = Nothing

End Sub