Attribute VB_Name = "All_Pic_GetText"
Private Function test_Get_Image_Str()

Debug.Print Get_Image_Str("C:\1\1.bmp")

End Function


Public Function Get_Image_Str(Path) As String
On Error GoTo Err1

    Dim midoc As Object
    
    Set midoc = CreateObject("MODI.Document")
    
    midoc.Create Path
    midoc.OCR MODI.miLANG_ENGLISH, True, True
    Get_Image_Str = midoc.images(0).Layout.Text
    Exit Function
    
Err1:
    
    MsgBox "执行程序不成功,请检查Microsoft Office Tools里面是否有Microsoft Office Document Imaging这个软件,如果没有,请向IT申请安装!", , "提示"
    
End Function

 
Public Function Image_Str(Path, Optional OneNote As Object) As String
 

Select Case VBA.Val(Application.Version)
    
    Case 12 '07版Excel 需使用Microsoft Office Document Imaging 来实现图片文字提取
        Image_Str = Get_Image_Str(Path)
        
    Case 14
        Image_Str = GetPictureText(Path, OneNote)
    
    Case Else
        
        MsgBox "本电脑Excel版本不能使用图片文字提取功能！"
        End
    
End Select


End Function
