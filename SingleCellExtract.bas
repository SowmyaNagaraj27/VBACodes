Attribute VB_Name = "Module2"
Option Compare Text
Public Function SingleCellExtract(ByVal LookupValue As String, LookupRange As Range, ColumnNumber As Integer, Char As String)
    Dim i As Long
    Dim xRet As String
    len1 = Len(LookupValue)
    For i = 1 To LookupRange.Columns(2).Cells.Count
        If Left(LookupRange.Cells(i, 1), len1) = LookupValue Then
            If xRet = "" Then
                xRetf = LookupRange.Cells(i, 1)  'firstname
                xRet = LookupRange.Cells(i, ColumnNumber) & Char 'orgname
                xRet1 = LookupRange.Cells(i, ColumnNumber - 1) 'last name
                xRet2 = LookupRange.Cells(i, ColumnNumber + 2) 'date mm/dd
                xRet3 = LookupRange.Cells(i, ColumnNumber + 3) 'year yyyy
'xRet = xRetf & "|" & xRet1 & "|" & xRet & xRet2 & "|" & xRet3 & "|"
                xRet = xRetf & "|" & xRet1 & "|" & xRet & xRet2 & "|" & xRet3 & "|"
            Else
                xRetf = LookupRange.Cells(i, 1)
                xRet1 = xRet & "" & xRetf & "|" & LookupRange.Cells(i, ColumnNumber - 1)
                xRet = LookupRange.Cells(i, ColumnNumber) & Char
                xRet2 = LookupRange.Cells(i, ColumnNumber + 2)
                xRet3 = LookupRange.Cells(i, ColumnNumber + 3)
'xRet = xRetf & "|" & xRet1 & "|" & xRet & xRet2 & "|" & xRet3 & "|"
'xRet = xRetf
                xRet = xRet1 & "|" & xRet & xRet2 & "|" & xRet3 & "|"
            End If
        End If
    Next
    SingleCellExtract = Left(xRet, Len(xRet))
End Function




