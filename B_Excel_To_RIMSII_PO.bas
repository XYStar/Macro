Attribute VB_Name = "B_Excel_To_RIMSII_PO"
 

Public Sub Excel_To_RIMSII_PO()

    If GetESC Then SetESC
    
    
Dim i As Integer
Dim j As Integer
Dim temp As String

Dim Endrow As Integer
Dim Erow As Integer
Dim Srow As Integer

    Endrow = Cells(60000, 1).End(xlUp).Row
    If Endrow < 12 Then
        ws_PopUp "没有数据!"
        Exit Sub
    End If
    
Dim OrderNO As String
Dim ItemID
Dim Vendor As String
Dim DueDate As String
Dim Comment As String

'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------
    
    Srow = VBA.Val(Cells(8, 2))
    Erow = VBA.Val(Cells(9, 2))
    
    If Srow = 0 Then
        Cells(8, 2) = 12
        Srow = 12
    End If
    
    If Erow = 0 Then
        Cells(9, 2) = Endrow
        Erow = Endrow
    End If
    
    If Application.Version = 14 Then Call CreateOneNote
    
    For i = Srow To Erow
        '----------
            OrderNO = Cells(i, Col("RVL SO NO"))
            Vendor = Cells(i, Col("Vendor"))
            ItemID = Cells(i, Col("Item"))
            
            PO_OrderNO OrderNO, ItemID
            PO_Vendor Vendor
            Save_Form_PO
            Cells(i, Col("RVL SO NO")).Interior.Color = 3407718
            Cells(i, Col("Item")).Interior.Color = 3407718
            Cells(i, Col("Vendor")).Interior.Color = 3407718
        '----------
            'PONO
            Cells(i, Col("PO#")) = Get_PONO '2016-08-18
        '----------
            Call PO_Comments(i)
            Cells(i, Col("Comment")).Interior.Color = 3407718
        '----------
        
        '----------
            DueDate = Cells(i, Col("Due Date"))
            Call PO_DueDate(DueDate)
            Cells(i, Col("Due Date")).Interior.Color = 3407718
        '----------
        
            Call PO_New_Orders
        
        
        Cells(i, Col("Status")) = "Done:" & Now
    Next i
    
    If Application.Version = 14 Then DeleteOneNotePages OneNote
        
    
End Sub
 
Private Function SetMyCol(Optional Dict As Dictionary, Optional myrow) As String()

    temp = "RVL SO No,Item,Vendor,Due Date,Comment,Status,PO#"
 
    MyCol = GetCol(temp, Dict, myrow)
        
End Function

Public Sub PO_Paste_ArrangeData()
    
    Sheets("PO Order form").Select
    Call PO_Paste_Data
    
    Sheets("Purchase Orders").Select
    ArrangeData
    
End Sub

Sub test()
    Date = "02/24/2016"
End Sub
Public Sub PO_ArrangeData()

Dim temp As String
Dim Data() As String
 
 
 temp = [B5]
 
' If temp = "" Then
'    ws_PopUp "请选择RBO!"
'    Exit Sub
' End If

'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------

 Select Case temp
    
    Case "Carter's", "Maurices"
 
            temp = Sheets("PO Order form").Cells(1, 1) & Sheets("PO Order form").Cells(1, 2)
            
            If temp = "" Then
                Sheets("PO Order form").Select
                ws_PopUp "没有找到数据!"
                Exit Sub
            End If
            '============
                
                GetData Col("RVL SO No"), "RVL SO No"
                GetData Col("Vendor"), "Vendor"
                GetData Col("Due Date"), "Due Date"
                GetData Col("Comment"), "Comment"
            
            '=============
            
    Case "Pumpkin Patch"
    
        Call SplictSheetData

            
    Case "Levis"
        Call Levis
        
End Select
 
        Erow = Cells(10000, 1).End(xlUp).Row
        Cells(8, 2) = 12
        Cells(9, 2) = Erow
                
        If Erow > 11 Then
                
            Range(Cells(12, 1), Cells(Erow, 20)).ClearComments
            Range(Cells(12, 1), Cells(Erow, 20)).Interior.Color = xlNone
        End If
        
        Cells(8, 4) = "整理完成"
        
End Sub
   
Private Function Levis()

    
    Dim i, j, k, t
    Dim temp
    Dim WK, TSH, SH, MSH
    Dim r1 As Range


'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------

    Set SH = ActiveSheet
    Set TSH = Nothing
    For Each WK In Workbooks
        For Each MSH In WK.Worksheets
            If MSH.Name = "Report" Then
                Set r1 = MSH.Cells.Find("PO Comment", , , xlWhole)
                If Not r1 Is Nothing Then
                    Set TSH = MSH
                    Exit For
                End If
            End If
        Next
        If Not TSH Is Nothing Then
            Exit For
        End If
    Next
    
    If TSH Is Nothing Then
        Exit Function
    End If
 
    Call PO_ClearData
    
    SH.Activate
    
    Set r1 = Cells.Find("*", , , , xlByRows, xlPrevious)
    Dim Erow, Srow, Tcol
    
    Erow = r1.Row
    
    If Erow > 5 Then
    Else
        Exit Function
    End If
    
    With TSH
        t = 12
        For i = 6 To Erow
            Tcol = .Cells.Find("SO#", , , xlWhole).Column
            If .Cells(i, Tcol) <> "" Then
                SH.Cells(t, Col("RVL SO No")) = .Cells(i, Tcol)
     
                SH.Cells(t, Col("Vendor")) = "*"
    
                Tcol = .Cells.Find("Due Date", , , xlWhole).Column
                SH.Cells(t, Col("Due Date")) = .Cells(i, Tcol).Text
    
                Tcol = .Cells.Find("PO Comment", , , xlWhole).Column
                SH.Cells(t, Col("Comment")) = .Cells(i, Tcol)
                t = t + 1
            End If
        Next i
    End With
 
End Function

   
Private Function SplictSheetData()
'2016-02-26
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim temp As String
Dim Data() As String
    
Dim r1 As Range
Dim SH As Worksheet
Dim MySH As Worksheet
    
    Set MySH = Sheets("Purchase Orders")
    Set SH = Sheets("PO Order form")
    
    'SH.Activate
    

Dim Endrow As Integer
Dim Srow As Integer
Dim Erow As Integer
Dim Trow As Integer
Dim Tcol As Integer
Dim MyCol As Integer
Dim myrow As Integer

Dim DueDate As String
Dim Vendor As String
Dim Comment As String
Dim CData() As String
Dim Sample As String
Dim StrData As String

'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------

    With SH
        '---------------
            'Vendor
            Set r1 = .Cells.Find("Vendor", , , xlPart)
            If Not r1 Is Nothing Then
                Data = VBA.Split(r1, ":")
                Vendor = "Avery Dennison (Guangzhou) Converted" '"Avery Dennison (Guangzhou) Converted Producted Ltd" 'Data(1)
                MyCol = r1.Column
            Else
                ws_PopUp "没有找到Vendor"
                Exit Function
            End If
            
            'DueDate
            Set r1 = .Cells.Find("Due Date", , , xlPart)
            If Not r1 Is Nothing Then
                Data = VBA.Split(r1, ":")
                DueDate = Data(1)
            Else
                ws_PopUp "没有找到Due Date"
                Exit Function
            End If
            
            'Comment
            Set r1 = .Cells.Find("comment", , , xlPart)
            If Not r1 Is Nothing Then
                Srow = r1.Row + 1
                t = 0
                For i = Srow To Srow + 10
                    temp = .Cells(i, MyCol)
                    temp = VBA.Trim(temp)
                    
                    If temp = "" Then
                        Exit For
                    End If
                    
                    ReDim Preserve CData(t)
                    temp = VBA.Replace(temp, VBA.ChrW(160), " ")
                    CData(t) = temp
                    t = t + 1
                Next i
            End If
        
        '---------------
        temp = "####################"
        Set r1 = .Cells.Find(temp, , , xlPart)
        If Not r1 Is Nothing Then
            myrow = r1.Row
            Tcol = r1.Column
            Erow = .Cells(50000, 1).End(xlUp).Row
            If Erow > myrow Then
                .Range(.Cells(myrow - 1, Tcol), .Cells(Erow, Tcol + 20)).ClearContents
            Else
                r1 = ""
            End If
        End If
        
        temp = "--------"
        Set r1 = .Cells.Find(temp, , , xlPart)
        If Not r1 Is Nothing Then
            MyCol = r1.Column
            myrow = r1.Row
        Else
            ws_PopUp "没有找到目标内容，请确认数据的正确性！"
            Exit Function
        End If
        
        Endrow = .Cells(60000, 1).End(xlUp).Row
        Trow = Endrow + 100
        .Cells(Trow, MyCol) = "####################"
        t = Trow + 1
        Srow = myrow + 1
        Erow = Endrow
        myrow = 12
        
        For i = Srow To Erow
            temp = .Cells(i, MyCol)
            temp = temp
            If InStr(1, temp, "---") Or VBA.Trim(temp) = "" Then
            Else
                
                Erase Data
                Call DataSplit(temp, Data)
                
                If UBound(Data) < 7 Then
                    ws_PopUp "Order form内容出错，请确保内容完整!"
                    End
                End If
                
                For j = 0 To UBound(Data)
                    .Cells(t, j + 1) = Data(j)
                Next j

                MySH.Cells(myrow, Col("RVL SO NO")) = Data(4) 'RVL SO NO
                MySH.Cells(myrow, Col("Vendor")) = Vendor
                MySH.Cells(myrow, Col("Due Date")) = DueDate
                    
                    Comment = ""
                    If UBound(Data) = 8 Then
                        temp = ""
                        Sample = Data(8)
                        For j = 0 To UBound(CData)
                            temp = CData(j)
                            If InStr(1, VBA.LCase(temp), "phx so") Then
                                temp = temp & Data(1)
                            End If
                            
                            If InStr(1, VBA.LCase(temp), "sample") Then
                                k = InStr(1, temp, ":")
                                StrData = VBA.Mid(temp, 1, k)
                                temp = StrData & Sample
                            End If
                            
                            Comment = Comment & temp & VBA.Chr(10)
                            
                        Next j
                    Else
                        
                        
                        '--------------------
                        Sample = "NO"
                        For j = 0 To UBound(CData)
                            temp = CData(j)
                            If InStr(1, VBA.LCase(temp), "phx so") Then
                                temp = temp & Data(1)
                            End If
                            
                            If InStr(1, VBA.LCase(temp), "sample") Then
                                If InStr(1, VBA.LCase(temp), "s#") Then
                                Else
                                    k = InStr(1, temp, ":")
                                    StrData = VBA.Mid(temp, 1, k)
                                    temp = StrData & Sample
                                End If
                            End If
                            
                            Comment = Comment & temp & VBA.Chr(10)
                        Next j
                        '--------------------
                    End If
                MySH.Cells(myrow, Col("Comment")) = Comment
                myrow = myrow + 1
                t = t + 1
            End If
        Next i
        
         '.Columns("B:Z").AutoFit

    End With
    
   
    
    
    
End Function

Private Function DataSplit(temp As String, Data() As String)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

Dim tempData As String
Dim StrData As String
Dim Tdata As String
'Dim temp As String
'Dim Data() As String
 
    
    tempData = VBA.Space(2)
    
    '========================
    
    t = 0
    Tdata = SpecialCharData(temp)
    
    i = InStr(1, Tdata, tempData)
    Do
        
        DoEvents
        StrData = VBA.Mid(Tdata, 1, i - 1)
        ReDim Preserve Data(t)
        Data(t) = StrData
        t = t + 1
        Tdata = VBA.Trim(VBA.Mid(Tdata, i)) & tempData
        If Tdata = tempData Then
            Exit Do
        End If
        i = InStr(1, Tdata, tempData)
    Loop
 
End Function

Private Function SpecialCharData(Tdata As String) As String
Dim i As Integer
Dim StrData As String

    '---------特殊字符替换为Space(1)
    For i = 1 To VBA.Len(Tdata)
        StrData = VBA.Mid(Tdata, i, 1)
        If VBA.AscW(StrData) > 122 Then
            Tdata = VBA.Replace(Tdata, StrData, " ")
        End If
    Next i
    '---------
    
    SpecialCharData = Tdata
    
End Function
Private Function GetData(MyCol As Integer, temp As String, Optional AfterData As String)
Dim i As Integer
Dim j As Integer

Dim Srow As Integer
Dim Erow As Integer

Dim r1 As Range
Dim r2 As Range

Dim Flag As Integer

    With Sheets("PO Order form")
        
        Erow = .Cells(50000, 1).End(xlUp).Row
        
        If AfterData = "" Then
            Set r1 = .Cells.Find(temp, , , xlWhole)
            If Not r1 Is Nothing Then
                Srow = r1.Row + 1
                j = 12
                For i = Srow To Erow
                    Cells(j, MyCol) = .Cells(i, r1.Column).Value
                    j = j + 1
                Next i
            Else
                .Select
                ws_PopUp "没有找到：" & temp & "列！"
                Exit Function
            End If
        Else
            Flag = 0
            Set r2 = .Cells.Find(AfterData, , , xlPart)
            If Not r2 Is Nothing Then
                Set r1 = .Cells.Find(temp, r2, , xlWhole)
                If Not r1 Is Nothing Then
                    Srow = r1.Row + 1
                    j = 12
                    For i = Srow To Erow
                        Cells(j, MyCol) = .Cells(i, r1.Column).Value
                        j = j + 1
                    Next i
                Else
                    Flag = 1
                End If
            Else
                Flag = 1
            End If
            
            If Flag = 1 Then
                .Select
                ws_PopUp "没有找到：" & temp & "列！"
                End
            End If
        End If
        
        
    End With
            
End Function


Public Sub PO_ClearData()
Dim Erow As Integer
Dim r1 As Range
    Set r1 = Cells.Find("*", , , , xlByRows, xlPrevious)
    Erow = r1.Row
    If Erow > 11 Then
        Range(Cells(12, 1), Cells(Erow, 10)).ClearContents
        Range(Cells(12, 1), Cells(Erow, 10)).Interior.Color = xlNone
    End If
    
    Cells(8, 2) = ""
    Cells(9, 2) = ""
    Cells(8, 4) = ""
    
    Rows("12:500").Delete
    '[B5] = ""
    
Dim DataSH As Worksheet

    Set DataSH = Sheets("PO Data Collection")
    
    With DataSH
        
        Set r1 = Cells.Find("*", , , , , xlPrevious)
        Erow = r1.Row
        
        If Erow > 1 Then
            .Range(.Cells(2, 1), .Cells(Erow, 10)).Interior.Color = xlNone
        End If
            
        .Cells.ClearContents
        
    End With
End Sub




Public Sub PO_ClearData_Data()
Dim Erow As Integer
    
    Cells.ClearContents
    Cells.ClearFormats
    
    
    Rows("20:200").Delete
    
    'Erow = Cells(10000, 1).End(xlUp).Row
    'Range(Cells(1, 1), Cells(Erow, 20)).ClearContents
    'Range(Cells(1, 1), Cells(Erow, 20)).ClearFormats '
    'Range(Cells(1, 1), Cells(Erow, 20)).Interior.Color = xlNone
    
End Sub

Public Sub PO_Paste_Data()
    Range("A1").Select
    ActiveSheet.Paste
End Sub


Public Sub Data_Collection()
    Select Case [B5]
        Case "Pumpkin Patch"
            Call Data_Collection_Pumpkin_Patch
        Case "Carter's"
            Call Data_Collection_Carters
        Case Else
            MsgBox "请选择RBO!"
    End Select
End Sub

Public Sub GetPO_From_Report()

Dim WK As Workbook
Dim SH As Worksheet
Dim MySH As Worksheet
Dim Flag As Integer
Dim i As Integer
Dim j As Integer
Dim t As Integer
Dim r1 As Range
Dim Data() As String

'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------


    Set MySH = ActiveSheet
    
    Flag = 0
    For Each WK In Workbooks
        For Each SH In WK.Worksheets
            If SH.Cells(1, 12) = "po_no" Then
                Flag = 1
                Exit For
            End If
        Next
        If Flag = 1 Then
            Exit For
        End If
    Next
    
    If Flag = 0 Then
        ws_PopUp "没有找到report表！"
        Exit Sub
    End If
    
Dim Erow As Integer
Dim temp As String
Dim Trow As Integer
    
    Erow = Cells(60000, 1).End(xlUp).Row
    
    If Erow < 12 Then
        ws_PopUp "没有找到内容！"
        Exit Sub
    End If
    
    Flag = 0
    For i = 12 To Erow
        temp = Cells(i, 1)
        Set r1 = SH.Columns(1).Find(temp, , , xlWhole)
        If Not r1 Is Nothing Then
            Cells(i, Col("PO#")) = SH.Cells(r1.Row, 12)
        Else
            Flag = 1
            Cells(i, Col("PO#")).Interior.Color = vbRed
        End If
    Next i
    
    If Flag = 1 Then
        Cells(8, 4) = "存在SO#没有找到PO#"
    Else
        Cells(8, 4) = "提取PO#完成"
    End If
    
    
End Sub

Private Sub Data_Collection_Pumpkin_Patch()

Dim i, j, k, t
Dim Data() As String
Dim Srow, Erow, Trow

    Dim TargetSH As Worksheet
    Dim DataSH As Worksheet
    Dim MySH As Worksheet


'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------


        Set MySH = ActiveSheet
        Erow = Cells(10000, 1).End(xlUp).Row
        If Erow < 12 Then
            Exit Sub
        End If
                
    Set TargetSH = Sheets("PO Order form")
    Set DataSH = Sheets("PO Data Collection")
    
    DataSH.Cells.ClearContents
    
        temp = "####################"
        Set r1 = TargetSH.Cells.Find(temp, , , xlWhole)
        If r1 Is Nothing Then
            ws_PopUp "没有找到关键内容!"
            End
        Else
            Trow = r1.Row + 1
        End If
    
    With DataSH
        Erow = TargetSH.Cells(50000, 1).End(xlUp).Row
        
        temp = "Phx PO#,Rims PO#,Rims Item#,LLKK or LB,Qty"
        Erase Data
        Data = VBA.Split(temp, ",")
        For i = 1 To 5
            .Cells(1, i) = Data(i - 1)
        Next i
        
        t = 2
        For i = Trow To Erow
            .Cells(t, 1) = TargetSH.Cells(i, 2).Value
            .Cells(t, 3) = TargetSH.Cells(i, 6).Value
            .Cells(t, 5) = TargetSH.Cells(i, 7).Value
            t = t + 1
        Next i

        t = 12
        For i = 2 To Erow
            .Cells(i, 2) = MySH.Cells(t, Col("PO#"))
            If MySH.Cells(t, Col("PO#")).Interior.Color <> vbWhite Then
                .Cells(i, 2).Interior.Color = MySH.Cells(t, Col("PO#")).Interior.Color
            End If
            t = t + 1
        Next i
    End With
    
    DataSH.Columns("A:Z").AutoFit
    
    DataSH.Select
    
End Sub

Private Sub Data_Collection_Carters()
Dim i, j, k, t
Dim Data() As String
Dim Srow, Erow, Trow

    Dim TargetSH As Worksheet
    Dim DataSH As Worksheet
    Dim MySH As Worksheet

'------
Dim Col As Dictionary
Dim ColData() As String
ColData = SetMyCol(Col, Trow)
'------

        Set MySH = ActiveSheet
        
        Erow = Cells(10000, 1).End(xlUp).Row
        If Erow < 12 Then
            Exit Sub
        End If
        
    Set TargetSH = Sheets("PO Order form")
    Set DataSH = Sheets("PO Data Collection")
    
    DataSH.Cells.ClearContents
    
    
    With DataSH
        Erow = TargetSH.Cells(50000, 1).End(xlUp).Row
        
        For i = 1 To Erow
            For j = 1 To 6
                .Cells(i, j) = TargetSH.Cells(i, j).Value
            Next j
        Next i
        
        t = 11
        For i = 1 To Erow
            .Cells(i, 7) = MySH.Cells(t, Col("PO#"))
            .Cells(i, 7).Interior.Color = MySH.Cells(t, Col("PO#")).Interior.Color
            t = t + 1
        Next i
    End With
    
    DataSH.Columns("A:Z").AutoFit
    
    DataSH.Select
    
End Sub

Public Sub Other_Tools()

Dim temp, Path
temp = Range("B5")

If temp = "" Then
    MsgBox "请选择RBO!"
    Exit Sub
End If

Select Case temp
    Case "Levis"
        Path = "\\fspdnan33\home$\CS\OM\Share\Project\ExceltoSystem\RIMSII\Levis  issue rims po.xlsm"
        Workbooks.Open Path, , True
End Select

End Sub
