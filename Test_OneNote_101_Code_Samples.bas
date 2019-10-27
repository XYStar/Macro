Attribute VB_Name = "Test_OneNote_101_Code_Samples"

' OneNote 2010

' Demonstrate the GetHierarchy method.

' Use any VBA host including Excel 2010, PowerPoint 2010,
' or Word 2010.
' OneNote 2010 is not a VBA host.

' In your VBA host, add references to the following
' external libraries using the Add References dialog:
' Microsoft OneNote 14.0 Object Library
' Microsoft XML, v6.0

' OneNote's GetHierarchy method allows you
' to get metadata and data about the OneNote
' Notebooks.

' Paste all this code into a module,
' place the cursor within the
' ListFivePagesFromFirstSectionOfFirstNotebook procedure,
' and press F5.
'
' The ListFivePagesFromFirstSectionOfFirstNotebook procedure
' uses the MSXML library to parse the returned XML
' from OneNote and output Notebook metadata
' to the Immediate window of your VBA host.

' In order to do this, the code loads up the list of available
' Notebooks and uses the first one found. It then gets the first
' Section from this Notebook.

Sub ListFivePagesFromFirstSectionOfFirstNotebook()
    ' Connect to OneNote 2010.
    ' OneNote will be started if it's not running.
    Dim OneNote As Object
    Set OneNote = CreateObject("OneNote.Application")
    
    ' Get all of the Notebook nodes.
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = GetFirstOneNoteNotebookNodes(OneNote)
    If Not nodes Is Nothing Then
        ' Get the first notebook found
        Dim node As MSXML2.IXMLDOMNode
        Set node = nodes(0)
        Dim noteBookName As String
        noteBookName = node.Attributes.getNamedItem("name").Text
        
        ' Get the ID so we can lookup the sections.
        Dim notebookID As String
        notebookID = node.Attributes.getNamedItem("ID").Text
        
        ' Load the XML for the sections for the notebook requested.
        Dim sectionsXml As String
        OneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
        
        Dim secDoc As MSXML2.DOMDocument
        Set secDoc = New MSXML2.DOMDocument
    
        If secDoc.LoadXML(sectionsXml) Then
            Dim secNodes As MSXML2.IXMLDOMNodeList
            Set secNodes = secDoc.DocumentElement.SelectNodes("//one:Section")

            If Not secNodes Is Nothing Then
                Dim secNode As MSXML2.IXMLDOMNode
                Set secNode = secNodes(0)
                
                Dim sectionName As String
                sectionName = secNode.Attributes.getNamedItem("name").Text
                 
                Dim sectionID As String
                sectionID = GetAttributeValueFromNode(secNode, "ID")
                                
                ' Load the XML for the Pages for the Section requested.
                Dim pagesXml As String
                OneNote.GetHierarchy sectionID, hsPages, pagesXml, xs2010
                
                Dim pagesDoc As MSXML2.DOMDocument
                Set pagesDoc = New MSXML2.DOMDocument
                
                If pagesDoc.LoadXML(pagesXml) Then
                    Dim pageNodes As MSXML2.IXMLDOMNodeList
                    Set pageNodes = pagesDoc.DocumentElement.SelectNodes("//one:Page")
                    
                    If Not pageNodes Is Nothing Then
                        ' Print out data about the Notebook, Section, and first
                        ' five pages of content.
                        Debug.Print "Notebook Name: " & noteBookName
                        Debug.Print "Notebook ID: " & notebookID
                        Debug.Print "  Section Name: " & sectionName
                        Debug.Print "  Section ID: " & sectionID
                        
                        ' Only show first five pages of information.
                        Const MAX_PAGES = 5
                        Dim intPageCount As Integer
                        intPageCount = 0
                        
                        Debug.Print "    *** Pages ***"
                        Dim pageNode As MSXML2.IXMLDOMNode
                        For Each pageNode In pageNodes
                            Debug.Print "    Page Name: " & GetAttributeValueFromNode(pageNode, "name")
                            Debug.Print "      ID: " & GetAttributeValueFromNode(pageNode, "ID")
                            Debug.Print "      Date Time: " & GetAttributeValueFromNode(pageNode, "dateTime")
                            Debug.Print "      Last Modified: " & GetAttributeValueFromNode(pageNode, "lastModifiedTime")
                            Debug.Print "      Page Level: " & GetAttributeValueFromNode(pageNode, "pageLevel")
                            intPageCount = intPageCount + 1
                            
                            If intPageCount = MAX_PAGES Then
                                Exit For
                            End If
                        Next
                    Else
                        MsgBox "OneNote 2010 Page nodes not found."
                    End If
                Else
                    MsgBox "OneNote 2010 Pages XML Data failed to load."
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
    
End Sub

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
    Dim doc As MSXML2.DOMDocument
    Set doc = New MSXML2.DOMDocument
    
    If doc.LoadXML(notebookXml) Then
        Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.SelectNodes("//one:Notebook")
    Else
        Set GetFirstOneNoteNotebookNodes = Nothing
    End If
End Function
