Attribute VB_Name = "A_Excel_To_RIMSII_SO"
Dim BeginRow As Integer
Dim EndCol As Integer
    
Public Sub Excel_Arrange_SO()
    Application.StatusBar = ""
    Dim r1 As Range, r2 As Range
    Dim Endrow As Integer
    Dim Col As Dictionary
    Dim ColData() As String
    Dim i%, j%, k%, t%
    Dim Srow%, Erow%, Scol%, Ecol%, Trow%, Tcol%
    Dim temp As String
    Dim StrData As String, Sdata As String
 
        
        ColData = MyCol(Col, Trow)
        BeginRow = Trow + 1
        
        EndCol = Cells(Trow, 256).End(xlToLeft).Column
        Erow = Cells(10000, 1).End(xlUp).Row
        If Erow > Trow Then
            Range(Cells(BeginRow, 1), Cells(Erow, 1)).ClearContents
        End If
        
        Set r1 = Cells.Find("*", , , , , xlPrevious)
        Endrow = r1.Row

        If Endrow < BeginRow Then
            ws_PopUp "没有找到内容！"
            End
        End If
        
        Erow = Endrow
        Ecol = EndCol
        Range(Cells(BeginRow, 1), Cells(Erow, Ecol)).Interior.Color = xlNone
        
        t = 1
        For i = BeginRow To Erow
            temp = VBA.Trim(Cells(i, Col("master customer")))
            If temp <> "" Then
                 Cells(i, 1) = VBA.Format(t, "0.0")
                 t = t + 1
            End If
        Next i
        
        For i = BeginRow To Erow
            temp = VBA.Trim(Cells(i, Col("Item NO.")))
            If temp <> "" Then
                StrData = VBA.Trim(Cells(i, 1))
                If StrData = "" Then
                    Cells(i, 1) = VBA.Val(Sdata) + 0.1
                    Sdata = Cells(i, 1)
                Else
                    Sdata = StrData
                End If
                
            End If
        Next i

        
        For i = BeginRow To Erow
            temp = VBA.Trim(Cells(i, Col("Bill to Customer")))
            If temp <> "" Then
                If InStr(temp, "-") Then
                    temp = VBA.Replace(temp, "-", "/")
                    Cells(i, Col("Bill to Customer")) = temp
                End If
            End If
            
            temp = VBA.Trim(Cells(i, Col("Ship to Address")))
            If temp <> "" Then
                If InStr(temp, "-") Then
                    temp = VBA.Replace(temp, "-", "/")
                    Cells(i, Col("Ship to Address")) = temp
                End If
            End If

            temp = VBA.Trim(Cells(i, Col("Order Recd Date")))
            If temp <> "" Then
                temp = VBA.Format(temp, "YYYY-MM-DD")
                Cells(i, Col("Order Recd Date")) = temp
            End If
            
            temp = VBA.Trim(Cells(i, Col("Customer Req Date")))
            If temp <> "" Then
                temp = VBA.Format(temp, "YYYY-MM-DD")
                Cells(i, Col("Customer Req Date")) = temp
            End If

            temp = VBA.Trim(Cells(i, Col("Promised Date")))
            If temp <> "" Then
                temp = VBA.Format(temp, "YYYY-MM-DD")
                Cells(i, Col("Promised Date")) = temp
            End If
        Next i
        
        
        For i = BeginRow To Erow
            temp = VBA.Trim(Cells(i, Col("QTY")))
            If temp = "" Then
                If r2 Is Nothing Then
                    Set r2 = Rows(i)
                End If
                Set r1 = Rows(i)
                Set r2 = Union(r2, r1)
            End If
        Next i
        
        If Not r2 Is Nothing Then
            r2.Delete
        End If

        Cells(BeginRow, 1).Select

        Cells(7, 3) = "1.0"
        Erow = Cells(10000, 1).End(xlUp).Row
        Cells(8, 3) = Cells(Erow, 1)
        
End Sub

Public Sub Clear_Data_SO()
    Dim r1 As Range
    Dim Erow As Integer
    
    Set r1 = Cells.Find("*", , , , , xlPrevious)
    
    Erow = r1.Row
    
    Dim Trow%
    Call MyCol(, Trow)
    BeginRow = Trow + 1
    EndCol = Cells(Trow, 256).End(xlToLeft).Column
    
    If Erow > Trow Then
        Range(Cells(BeginRow, 1), Cells(Erow, EndCol)).Interior.Color = xlNone
        Range(Cells(BeginRow, 1), Cells(Erow, EndCol)).ClearContents
        Range(Cells(BeginRow, 1), Cells(Erow, EndCol)).ClearComments
    End If
    
    Rows(BeginRow & ":500").Delete
    
    Cells(5, 3) = "New"
    Cells(6, 3) = "F8"
    Cells(7, 3) = ""
    Cells(8, 3) = ""
    
    Cells(BeginRow, 1).Select
    
End Sub

Public Sub Excel_To_RIMSII_SO()
    Application.StatusBar = ""
    
    Dim i%, j%, k%, t%
    Dim Srow%, Erow%, Endrow%, Trow%, Scol%, Ecol%, Tcol%, myrow%, SubSrow%, SubErow%
    Dim temp As String
    Dim Data() As String
    Dim Tdata
    Dim r1 As Range
    Dim Col As Dictionary
    Dim Dict As Dictionary
    Dim ColData() As String
    Dim MyStep As String
    Dim hwnd As Long
           
        If Application.Version = 14 Then Call CreateOneNote
 
              
        If GetESC Then SetESC
        Call Init
        hwnd = setHwnd
        
        ColData = MyCol(Col, Trow)
        BeginRow = Trow + 1
        
        EndCol = Cells(Trow, 256).End(xlToLeft).Column
        
        Set r1 = Cells.Find("*", , , , , xlPrevious)
        If Not r1 Is Nothing Then
            Endrow = r1.Row
        Else
            ws_PopUp "没有找到最后一行"
            End
        End If
        
        If Endrow < BeginRow Then
            ws_PopUp "没有找到内容!"
            End
        End If
        
        Trow = Cells(65535, 1).End(xlUp).Row
        If Trow > BeginRow - 1 Then
            For i = BeginRow To Endrow
                temp = VBA.Trim(Cells(i, Col("master customer")))
                If temp <> "" Then
                    If Cells(i, 1) = "" Then
                        ws_PopUp "请先整理数据!"
                        End
                    End If
                End If
                temp = VBA.Trim(Cells(i, Col("Item NO.")))
                If temp <> "" Then
                    If Cells(i, 1) = "" Then
                        ws_PopUp "请先整理数据!"
                        End
                    End If
                End If
            Next i
        Else
            ws_PopUp "请先整理数据!"
            End
        End If
        
        Ecol = EndCol
        Srow = BeginRow
        Erow = Endrow
        '=========
        Dim FromNO As String
        Dim ToNO As String
            FromNO = Cells(7, 3)
            ToNO = Cells(8, 3)
            If FromNO <> "" Then
                Set r1 = Columns(1).Find(FromNO, , , xlWhole)
                If r1 Is Nothing Then
                    ws_PopUp "没有在导单表中找到 From NO.: " & FromNO
                    End
                End If
                Srow = r1.Row
            End If
            
            If ToNO <> "" Then
                Set r1 = Columns(1).Find(ToNO, , , xlWhole)
                If r1 Is Nothing Then
                    ws_PopUp "没有在导单表中找到 To NO.: " & ToNO
                    End
                End If
                Trow = r1.Row
                For j = Trow + 1 To Erow
                    If Cells(j, 1) <> "" Then
                        Trow = j - 1
                        Exit For
                    End If
                    If j = Erow Then
                        Trow = j
                    End If
                Next j
                Erow = Trow
            End If
        '=========
            
        'Range(Cells(BeginRow, 1), Cells(Erow, Ecol)).Interior.Color = xlNone
        
        '----------- 0804
        'New 新建 F3,F4,F5,F7,F8如文字描述
        Dim BeginStr As String
        Dim EndStr As String
        Dim StepBegin As Integer
        Dim StepEnd As Integer
        
            BeginStr = Cells(5, 3)
            EndStr = Cells(6, 3)
            If BeginStr = "" Then BeginStr = "New"
            If EndStr = "" Then EndStr = "F8"
            Tdata = Array("New", "F3", "F4", "F5", "F7", "F8")
            For i = 0 To UBound(Tdata)
                If Tdata(i) = BeginStr Then
                    StepBegin = i
                End If
                
                If Tdata(i) = EndStr Then
                    StepEnd = i
                End If
            Next i
            
            If StepBegin > StepEnd Then
                ws_PopUp "导单表格步骤设置错误,请重新设置!"
                End
            End If
        '-----------
        
        For i = Srow To Erow
            If Cells(i, 1) <> "" Then
                myrow = i
                Trow = myrow
                '----
                For j = myrow + 1 To Erow
                    If Cells(j, 1) <> "" Then
                        Trow = j - 1
                        Exit For
                    End If
                    If j = Erow Then
                        Trow = j
                    End If
                Next j
                SubSrow = myrow
                SubErow = Trow
                
                For j = StepBegin To StepEnd
                
                    MyStep = Tdata(j)
                    Select Case MyStep
                        Case "New"
                            Call To_RIMSII_Step_New(myrow) 'New
                        Case "F3"
                            Call To_RIMSII_Step_F3(myrow) 'F3
                        Case "F4"
                            Call To_RIMSII_Step_F4(myrow) 'F4
                        Case "F5"
                            Call To_RIMSII_Step_F5(myrow) 'F5
                        Case "F7"
                            Call To_RIMSII_Step_F7(myrow, SubSrow, SubErow) 'F7
                        Case "F8"
                            Call To_RIMSII_Step_F8(myrow) 'F8
                    End Select
                    
                Next j
                    
                Delay 1000
                
                Cells(myrow, Col("Record Time")) = "Done " & Now
                Cells(myrow, Col("Record Time")).Interior.Color = vbYellow
            End If
        Next i
        
        If Application.Version = 14 Then DeleteOneNotePages OneNote
        
    ws_PopUp "完成！"
    
End Sub
Private Function To_RIMSII_Step_New(myrow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim temp As String

    If GetESC Then SetESC
    Call Init
    
    ColData = MyCol(Col)
    
        Dim Master$
        temp = Cells(myrow, Col("master customer"))
        Master = VBA.Trim(temp)
        If Master <> "" And Cells(myrow, Col("master customer")).Interior.Color = vbWhite Then
            SendMyData "^n" 'Ctrl+n
            Delay 1000
        End If
            
End Function

Private Function To_RIMSII_Step_F3(myrow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim temp As String
'Dim MyRow As Integer
'MyRow = 12
     
    If GetESC Then SetESC
    Call Init
    
    ColData = MyCol(Col)
        
        Dim Master$
        temp = Cells(myrow, Col("master customer"))
        Master = VBA.Trim(temp)

        Dim Billto$
        temp = Cells(myrow, Col("bill to customer"))
        Billto = VBA.Trim(temp)

        Dim PO$
        temp = Cells(myrow, Col("customer po no."))
        PO = VBA.Trim(temp)
        
        Dim OrderRecdDate$
        temp = Cells(myrow, Col("Order Recd Date"))
        OrderRecdDate = VBA.Trim(temp)
        
        Dim OrderNO$
        OrderNO = ""
        
        '=======
            If Cells(myrow, Col("master customer")).Interior.Color = vbWhite Then
                Call Step_F3(Master, Billto, PO, OrderRecdDate, OrderNO, myrow, Col)
            Else
                Exit Function
            End If
        '=======
        
        '记录Order NO.
        Cells(myrow, Col("Order NO.")) = ""
        If OrderNO <> "" Then
            Cells(myrow, Col("Order NO.")) = OrderNO
            Cells(myrow, Col("Order NO.")).Interior.Color = vbYellow
        End If
End Function

Private Function To_RIMSII_Step_F4(myrow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim temp As String
'Dim MyRow As Integer
'MyRow = 12

        If GetESC Then SetESC
        Call Init
        
        ColData = MyCol(Col)
        
        Dim Shipto$
        temp = Cells(myrow, Col("ship to address"))
        Shipto = VBA.Trim(temp)
        
        '====
            If Cells(myrow, Col("ship to address")).Interior.Color = vbWhite Then
                Call Step_F4(Shipto, myrow, Col)
            Else
                Exit Function
            End If
        '====

        
End Function

Private Function To_RIMSII_Step_F5(myrow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim temp As String
'Dim MyRow As Integer
'MyRow = 12

        If GetESC Then SetESC
        Call Init
        
        ColData = MyCol(Col)
        
        Dim ItemNO$
        temp = Cells(myrow, Col("Item NO."))
        ItemNO = VBA.Trim(temp)

        Dim Total$
        temp = Cells(myrow, Col("Total"))
        Total = VBA.Trim(temp)

        Dim Price$
        temp = Cells(myrow, Col("Unit Price"))
        Price = VBA.Trim(temp)

        Dim StyleNO$
        temp = Cells(myrow, Col("Style NO."))
        StyleNO = VBA.Trim(temp)
        
        '=====
            If Cells(myrow, Col("Item NO.")).Interior.Color = vbWhite Then
                Call Step_F5(ItemNO, Total, Price, StyleNO, myrow, Col)
            Else
                Exit Function
            End If
        '=====

        
End Function

Private Function To_RIMSII_Step_F7(myrow As Integer, SubSrow As Integer, SubErow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim Dict As Dictionary
Dim MP As Dictionary
Dim temp As String
Dim i As Integer, j As Integer
'Dim MyRow As Integer
'MyRow = 12
'Dim SubSrow As Integer, SubErow As Integer
'SubSrow = 12
'SubErow = 12

        Set MP = CreateObject("Scripting.Dictionary")
        
        If GetESC Then SetESC
        Call Init
        
        ColData = MyCol(Col)
        
            Dim Size$, QTY$
            Set Dict = Nothing
            For j = SubSrow To SubErow
                temp = Cells(j, Col("Size"))
                Size = VBA.Trim(temp)
                
                temp = Cells(j, Col("QTY"))
                QTY = VBA.Trim(temp)

                If Size <> "" And _
                   QTY <> "" And _
                   Cells(j, Col("Size")).Interior.Color = vbWhite And _
                   Cells(j, Col("QTY")).Interior.Color = vbWhite Then
                    '-----
                        Set Dict = Set_Dict(Size, QTY, Dict)
                        MP(j) = Col("Size")
                    '-----
                End If
            Next j
            
            Call Step_F7(Dict, MP)
            
End Function

Private Function To_RIMSII_Step_F8(myrow As Integer)
Dim Col As Dictionary
Dim ColData() As String
Dim temp As String
'Dim MyRow As Integer
'MyRow = 12

        If GetESC Then SetESC
        Call Init
        
        ColData = MyCol(Col)
        
        Dim ReqDate$, PromisedDate$
        temp = Cells(myrow, Col("customer req date"))
        ReqDate = VBA.Trim(temp)


        temp = Cells(myrow, Col("promised date"))
        PromisedDate = VBA.Trim(temp)

        If Cells(myrow, Col("customer req date")).Interior.Color = vbWhite Then
            Call Step_F8(ReqDate, PromisedDate, myrow, Col)
        End If
        
End Function


Private Function MyCol(Optional Dict As Dictionary, Optional myrow) As String()
Dim temp As String
    temp = "Master Customer,Bill to Customer,Customer PO NO.,Order Recd Date,Ship to Address,Item NO.,Total,Unit Price,Style NO.,Size,QTY,Customer Req Date,Promised Date,Order NO.,Record Time"
    MyCol = GetCol(temp, Dict, myrow)
        
End Function

Public Sub Doc_Show_SO()
On Error Resume Next
    
    Dim Wd
    
    Set Wd = CreateObject("word.application")
    Wd.documents.Open "\\fspdnan33\home$\CS\OM\Share\Project\ExceltoSystem\RIMSII\RIMSII SO导单注意事项.docx", ReadOnly:=True
    Wd.Visible = True
End Sub

 
 
