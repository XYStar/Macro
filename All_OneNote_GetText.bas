Attribute VB_Name = "All_OneNote_GetText"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub LoadFile()
 Dim cht As Chart
 Dim Obj As Object
 Dim shp As Shape
 
 
 Set cht = ThisWorkbook.Charts("Chart1")
 cht.Select
 ActiveSheet.Paste
 Set shp = cht.Shapes(1)

'Exporting the Chart to Image file.
 cht.Export ThisWorkbook.Path & "Test1.bmp"

'Calling the Second method to read text
 GetPictureText ThisWorkbook.Path & "Test1.bmp"

'Deleting the image file
 Kill ThisWorkbook.Path & "Test1.bmp"
 
'Deleting the image from the Chart for future use
 cht.Shapes.Range(Array("chart")).Delete
End Sub


Private Function Test_Create_OneNote()
    
    Dim OneNote As Object
    Call Kill_OneNote
    Set OneNote = CreateObject("OneNote.Application")
    
    For i = 1 To 3
        Dim strfile As String
        strfile = "C:\1\1.bmp"
        Debug.Print GetPictureText(strfile, OneNote)
    Next
    
End Function
Public Function CreateOneNote()

    If Application.Version = 14 Then
        If OneNote Is Nothing Then
            Set OneNote = CreateObject("OneNote.Application")
        End If
    End If
    
End Function
Public Function Kill_OneNote()
    On Error Resume Next
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='OneNote.exe'")
        Process.Terminate (0)
        Exit For
    Next
    Sleep 200
End Function

Private Sub test_GetPictureText()
    Dim Path As String
        Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For PO Release\1.bmp"
        Path = "C:\1\1.bmp"
        
        Set OneNote = CreateObject("OneNote.Application")
        
        Debug.Print GetPictureText(Path, OneNote)
        If Application.Version = 14 Then DeleteOneNotePages OneNote
        
End Sub
'2016-07-15 Star He
Public Function GetPictureText(strfile, Optional OneNote As Object) As String
 
    Dim bytFile() As Byte
    Dim base64String, imageXmlStr
    Dim nCounter As Integer
    Dim Bitmap


    If OneNote Is Nothing Then
        Set OneNote = CreateObject("OneNote.Application")
    End If
              
    ' Get all of the Notebook nodes.
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = GetFirstOneNoteNotebookNodes(OneNote)
 
    Set Bitmap = LoadPicture(strfile)
    bytFile = GetFileBytes(strfile)
    base64String = EncodeBase64(bytFile)
    
    Dim ww, hh
    ww = Round(Bitmap.Width / 10)
    hh = Round(Bitmap.Height / 10)
 
    
    If Not nodes Is Nothing Then
        ' Get the first OneNote Notebook in the XML document.
        Dim node As MSXML2.IXMLDOMNode
        Set node = nodes(0)
        Dim noteBookName As String
        noteBookName = node.Attributes.getNamedItem("name").Text
         
        ' Get the ID for the Notebook so the code can retrieve
        ' the list of sections.
        Dim notebookID As String
        notebookID = node.Attributes.getNamedItem("ID").Text
         
        ' Load the XML for the Sections for the Notebook requested.
        Dim sectionsXml As String
        OneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
                  
        Dim secDoc As MSXML2.DOMDocument60
        Set secDoc = New MSXML2.DOMDocument60
     
        If secDoc.LoadXML(sectionsXml) Then
            ' select the Section nodes
            Dim secNodes As MSXML2.IXMLDOMNodeList
            Set secNodes = secDoc.DocumentElement.getElementsByTagName("one:Section")
                        
            If Not secNodes Is Nothing Then
                ' Get the first section.
                Dim secNode As MSXML2.IXMLDOMNode
                Set secNode = secNodes(0)
 
                Dim sectionID As String
                sectionID = secNode.Attributes.getNamedItem("ID").Text

                ' Create a new blank Page in the first Section
                ' using the default format.
                Dim newPageID As String

                OneNote.CreateNewPage sectionID, newPageID, npsDefault
                
                ' Get the contents of the page.
                Dim outXML As String
                
                OneNote.GetPageContent newPageID, outXML, piAll, xs2010
                 
                Dim doc As MSXML2.DOMDocument60
                Set doc = New MSXML2.DOMDocument60
                ' Load Page's XML into a MSXML2.DOMDocument60 object.
                If doc.LoadXML(outXML) Then
                    ' Get Page Node.
                    Dim pageNode As MSXML2.IXMLDOMNode
                    Set pageNode = doc.getElementsByTagName("one:Page")(0)
                                        
                    Dim newElement As MSXML2.IXMLDOMElement
                    Dim newNode As MSXML2.IXMLDOMNode
                     
                    ' Create Outline node.
                    Set newElement = doc.createElement("one:Outline")
                    newElement.setAttribute "lang", "en-US"
                    Set newNode = pageNode.appendChild(newElement)
                    ' Create OEChildren.
                    Set newElement = doc.createElement("one:OEChildren")
                    Set newNode = newNode.appendChild(newElement)
                    ' Create OE.
                    Set newElement = doc.createElement("one:OE")
                    newElement.setAttribute "lang", "en-US"
                    Set newNode = newNode.appendChild(newElement)
                    
                    ' Create Image.
                    Set newElement = doc.createElement("one:Image")
                    'newElement.setAttribute "format", "bmp"
                    Set newNode = newNode.appendChild(newElement)
                    
                    ' Create Size.
                    Set newElement = doc.createElement("one:Size")
                    newElement.setAttribute "width", ww
                    newElement.setAttribute "height", hh
                    newElement.setAttribute "isSetByUser", "true"
                    newNode.appendChild newElement
                    
                    'Push the image bnary data
                    Set newElement = doc.createElement("one:Data")
                    newElement.Text = base64String
                    newNode.appendChild newElement
                  
                    ' Update OneNote with the new content.
                    OneNote.UpdatePageContent doc.XML, , , True
                   
                    Dim strxml As String
                    'Get the contnt back from OneNote Page
                    OneNote.GetPageContent newPageID, strxml
                    doc.LoadXML strxml
                    Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                    
                    nCounter = 1
                    
                    Do While nodes.Length = 0
                         OneNote.GetPageContent newPageID, strxml
                         doc.LoadXML strxml
                         Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                         nCounter = nCounter + 1
                         If nCounter = 100 Then
                             GetPictureText = "Image is not readable"
                             'OneNote.DeleteHierarchy newPageID '删除新建page
                             Exit Function
                         End If
                    Loop
                    
                    Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                    GetPictureText = nodes(0).Text '返回值 2016-07-14
                    'OneNote.DeleteHierarchy newPageID '删除新建page
                    
                End If
            Else
                MsgBox "OneNote 2010 Section nodes not found."
            End If
        Else
            MsgBox "OneNote 2010 Section XML Data failed to load."
        End If
    Else
        MsgBox "OneNote 2010 XML Data failed to load."
    End If
     
End Function

'2016-09-13
Public Function DeleteOneNotePages(OneNote As Object)
  
    If OneNote Is Nothing Then
        Exit Function
    End If
  
    ' Get all of the Notebook nodes.
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = GetFirstOneNoteNotebookNodes(OneNote)

    If Not nodes Is Nothing Then
        ' Get the first OneNote Notebook in the XML document.
        Dim node As MSXML2.IXMLDOMNode
        Set node = nodes(0)

        ' Get the ID for the Notebook so the code can retrieve
        ' the list of sections.
        Dim notebookID As String
        notebookID = node.Attributes.getNamedItem("ID").Text
         
        ' Load the XML for the Sections for the Notebook requested.
        Dim sectionsXml As String
        OneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
                                 
        Dim secDoc As MSXML2.DOMDocument60
        Set secDoc = New MSXML2.DOMDocument60
     
        If secDoc.LoadXML(sectionsXml) Then
            ' select the Section nodes
            Dim secNodes As MSXML2.IXMLDOMNodeList
            Set secNodes = secDoc.DocumentElement.getElementsByTagName("one:Section")

            If Not secNodes Is Nothing Then
                ' Get the first section.
                Dim secNode As MSXML2.IXMLDOMNode
                Set secNode = secNodes(0)

                Dim sectionID As String
                sectionID = secNode.Attributes.getNamedItem("ID").Text

                '=========
                ' Load the XML for the Pages for the Section requested.
                Dim pagesXml As String
                OneNote.GetHierarchy sectionID, hsPages, pagesXml, xs2010
                
                Dim pagesDoc As MSXML2.DOMDocument
                Set pagesDoc = New MSXML2.DOMDocument
                If pagesDoc.LoadXML(pagesXml) Then
                    Dim pageNodes As MSXML2.IXMLDOMNodeList
                    Set pageNodes = pagesDoc.DocumentElement.SelectNodes("one:Page")
                
                    If Not pageNodes Is Nothing Then
                        Dim pageNode As MSXML2.IXMLDOMNode
                        For Each pageNode In pageNodes
                            OneNote.DeleteHierarchy GetAttributeValueFromNode(pageNode, "ID")
                            DoEvents
                        Next
                    End If
                End If
                '=========
            End If 'If Not secNodes Is Nothing
        End If 'f secDoc.LoadXML(sectionsXml)
    End If 'If Not nodes Is Nothing

End Function


Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String
    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        GetAttributeValueFromNode = "Not found."
    Else
        GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
    End If
End Function

Private Function GetFirstOneNoteNotebookNodes(OneNote As Object) As MSXML2.IXMLDOMNodeList
    ' Get the XML that represents the OneNote notebooks available.
    Dim notebookXml As String
    ' OneNote fills notebookXml with an XML document providing information
    ' about what OneNote notebooks are available.
    ' You want all the data and thus are providing an empty string
    ' for the bstrStartNodeID parameter.
    OneNote.GetHierarchy "", hsNotebooks, notebookXml, xs2010
     
    ' Use the MSXML Library to parse the XML.
    Dim doc As MSXML2.DOMDocument60
    Set doc = New MSXML2.DOMDocument60
     
    If doc.LoadXML(notebookXml) Then
        Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.getElementsByTagName("one:Notebook")
    Else
        Set GetFirstOneNoteNotebookNodes = Nothing
    End If
End Function

Private Function GetFileBytes(strPath) As Byte()
    With CreateObject("ADODB.Stream")
        .Open
        .type = 1  ' adTypeBinary
        .LoadFromFile strPath
        GetFileBytes = .Read
        .Close
    End With
    
End Function

Private Function EncodeBase64(arrData() As Byte) As String
  'arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.Text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

