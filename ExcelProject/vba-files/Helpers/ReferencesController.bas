Attribute VB_Name = "ReferencesController"


'namespace=vba-files\Helpers


'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Sub ExcecuteRefLink()

    Dim cell As Range
    Dim range As Range
    Dim firstAddress As String
    Dim searchValue As String
    Dim refvalue As String

    searchValue = "img="
    Set range = ActiveSheet.Cells

    With range
        Set cell = .Find(what:=searchValue, LookAt:=XlPart)
        If Not cell Is Nothing Then
            firstAddress = cell.Address
            Do
                If InStr(1, cell.Value, "img=") = 1 Then

                    refvalue = Mid(cell.Value, InStr(cell.Value, "img=") + 4)
                    If trim(refvalue & vbnullstring) = vbnullstring Then 
                        cell.Value = "!Error:Empty reference"
                    Else 
                        Call SearchAndInsertImage(cell, refvalue) 
                    End If 
                End If


                Set cell = .Find(what:=searchValue, LookAt:=XlPart)
            Loop While Not cell Is Nothing
        End If
    End With

    'vidage des variables
    Set range = Nothing
    Set cell = Nothing

End Sub


Sub SearchAndInsertImage(refcell As Range, refvalue as String) 

    Const factor = 0.9  'picture is 90% of the size of cell

    Dim cell As Range
    Dim cellPath as Range
    Dim range As Range
    Dim address As String

    Set range = ActiveWorkbook.Worksheets("Catalog").Columns("C")
    With range
        Set cell = .Find(what:=refvalue, LookAt:=xlWhole)
        If Not cell Is Nothing Then
            address = cell.Address
            refcell.Value = ""

            Set cellPath = cell.Offset(0, 1)
            Set p = ActiveSheet.Shapes.AddPicture(Filename:=cellPath.Value, LinkToFile:=False, SaveWithDocument:=True, _
                Left:=refcell.Left, Top:=refcell.Left, Width:=-1, Height:=-1)

            p.Width = refcell.Width * factor
            
            'adjust row height
            If refcell.RowHeight < p.Height / factor Then
                refcell.RowHeight = p.Height / factor
            End If

            p.Left = refcell.Left + (refcell.Width - p.Width) / 2
            p.Top = refcell.Top + (refcell.Height - p.Height) / 2

        Else
            refcell.Value = "!Error:Reference not found in catalog"
        End If
    End With

End Sub

