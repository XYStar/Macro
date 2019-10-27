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
    
    MsgBox "ִ�г��򲻳ɹ�,����Microsoft Office Tools�����Ƿ���Microsoft Office Document Imaging������,���û��,����IT���밲װ!", , "��ʾ"
    
End Function

 
Public Function Image_Str(Path, Optional OneNote As Object) As String
 

Select Case VBA.Val(Application.Version)
    
    Case 12 '07��Excel ��ʹ��Microsoft Office Document Imaging ��ʵ��ͼƬ������ȡ
        Image_Str = Get_Image_Str(Path)
        
    Case 14
        Image_Str = GetPictureText(Path, OneNote)
    
    Case Else
        
        MsgBox "������Excel�汾����ʹ��ͼƬ������ȡ���ܣ�"
        End
    
End Select


End Function
