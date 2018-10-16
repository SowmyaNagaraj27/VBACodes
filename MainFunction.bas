Attribute VB_Name = "Module1"
Private Sub Auto_open()
    Sheets("Sheet1").Select
    If Worksheets("sheet1").Range("A1").Value = "TRAINING LOCATION" Then
        Call Swap_Column
        Sheets("Sheet3").Select
        Worksheets("Sheet3").RefreshAll
    Else
        Sheets("Sheet3").Select
    End If
    Worksheets("Sheet3").Range("L2:N400").Clear
End Sub


Sub OrgName()
    Dim OrgName As Variant
    Dim fname As String
    Dim lname As Variant
    Dim sarray() As String
    Dim orgarray() As String
    
    Worksheets("Sheet3").Activate
    fname = Range("I10").Value
'MsgBox fname
    lname = Range("I11").Value
    
    Set box = Range("L2")
    
    
'Search based on first name
    If Len(fname) > 0 Then
        Worksheets("Sheet3").Range("L2:N400").Clear
        Application.ScreenUpdating = False
        If Worksheets("sheet1").Range("A1").Value = "Last" Then
            Worksheets("sheet1").Activate
            Call LastNameInterchange
            Sheets("Sheet3").Select
        End If
        OrgName = SingleCellExtract(fname, Sheet1.Range("A:E"), 3, "|")
        If Len(OrgName) = 0 Then
            GoTo handler
        End If
        sarray = Split(OrgName, "|")
        For i = LBound(sarray) To UBound(sarray)
            fname = sarray(i)
            i = i + 1
            lname = sarray(i)
            i = i + 1
'MsgBox "orgname" & sarray(i)
            On Error GoTo handler1
            Sal = Application.WorksheetFunction.VLookup(sarray(i), Sheet2.Range("A:B"), 1, False)
            i = i + 1
            cerdate = sarray(i)
            cerdate = Format(cerdate, "mm/dd")
'MsgBox "certdate" & sarray(i)
            i = i + 1
            year1 = sarray(i)
            year1 = cerdate & "/" & year1
            fname = WorksheetFunction.Proper(fname)
            box.Value = fname & " " & lname
            Set box = box.Offset(0, 1)
            box.Value = Sal
            Set box = box.Offset(0, 1)
            box.Value = year1
            Range("I10").ClearContents
label1:
            Set box = box.Offset(1, -2)
            Next i
'Search based on Lname
        ElseIf Len(lname) > 0 Then
            Worksheets("Sheet3").Range("L2:N400").Clear
            Application.ScreenUpdating = False
            If Worksheets("sheet1").Range("A1").Value = "NAME" Then
                Worksheets("sheet1").Activate
                Call LastNameInterchange
                Sheets("Sheet3").Select
            End If
            OrgName = SingleCellExtract(lname, Sheet1.Range("A:E"), 3, "|")
            sarray = Split(OrgName, "|")
            For i = LBound(sarray) To UBound(sarray)
                lname = sarray(i)
                i = i + 1
                fname = sarray(i)
                i = i + 1
                On Error GoTo handler2
                Sal = Application.WorksheetFunction.VLookup(sarray(i), Sheet2.Range("A:B"), 1, False)
                i = i + 1
                cerdate = sarray(i)
                cerdate = Format(cerdate, "mm/dd")
                'MsgBox "certdate" & sarray(i)
                i = i + 1
                year1 = sarray(i)
                year1 = cerdate & "/" & year1
                lname = WorksheetFunction.Proper(lname)
                box.Value = fname & " " & lname
                Set box = box.Offset(0, 1)
                box.Value = Sal
                Set box = box.Offset(0, 1)
                box.Value = year1
                Range("I11").ClearContents
label2:
                Set box = box.Offset(1, -2)
            Next i
        End If
    End
handler:
            MsgBox " Please Enter a valid First name or last name"
            Exit Sub
            
handler1:
            If Err.Number = 1004 Then
                fname = WorksheetFunction.Proper(fname)
                'MsgBox fname & " " & lname & " does not have a license"
                box.Value = fname & " " & lname
                Set box = box.Offset(0, 1)
                box.Value = "No Licence"
                Set box = box.Offset(0, 1)
                box.Value = "No Licence"
                i = i + 2
                Resume label1
            End If
            
handler2:
            If Err.Number = 1004 Then
                lname = WorksheetFunction.Proper(lname)
                'MsgBox fname & " " & lname & " does not have a license"
                box.Value = fname & " " & lname
                Set box = box.Offset(0, 1)
                box.Value = "No Licence"
                Set box = box.Offset(0, 1)
                box.Value = "No Licence"
                i = i + 2
                Resume label2
            End If
            
            
        End Sub
        

