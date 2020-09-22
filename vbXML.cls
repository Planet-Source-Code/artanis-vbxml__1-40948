VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
' vbXML : a set of functions designed to easily
' access the MSXML control provided by microsoft.
' vbXML may also be known as a class wrapper, but it
' is packed in to a DLL (Dynamically Linked Library)
' vbXML attemps to re-create the MSXML interface by
' Providing a set of front-end methods and properties
' that access MSXML directly from the MSXML DLL file.
' Uses for vbXML could be easy skins, personalized
' settings (for forms), and anything else you could
' use the MSXML control for.

' See also: whats_new.txt, to_do.txt

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' vbXML is based entirely on the CGoXML
' wrapper class submited to PlanetSourceCode by
' Roman Kehr
'
' I have made plenty of fixes and added many new
' features (or I plan to add new features), but I
' thought you must know about CGoXML, which greatly
' inspired vbXML
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' In order to utilize most (if not all) of the
' functions included here you must understand
' (at least moderatley) XPath.
' What is XPath?
'   XPath is used by MSXML to access specific nodes
'   of an XML file
' How do I access a node?
'   XPath (the way I use it) holds this format:
'       "/parent1/child1/child2/childN/node"
'       Where N is any number
'   So to access a node, use this format:
'       "/example/child/node"
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' Any questions, comments, suggestions may be
' sent to:
'               akartanis@akguild.com
'---------------------------------------------------

Dim xDoc As MSXML.DOMDocument

Public Enum OpenXMLDoc
    oxFile = 0
    oxString = 1
End Enum

Public Enum DocInfoConst
    diVERSION = 0
    diENCODING = 1
    diSTANDALONE = 2
End Enum
'---------------------------------------------------
' Public Property Get
'   XMLDocumentInfo
' Parameters
'   diDesiredInfo As DocInfoConst : Tells the property
'   what to look for
' What it does
'   Gets the document info from the <?xml?> line
' Author
'   Anthony (11.19.02)
' Modified
'   (11.19.02)
'---------------------------------------------------
Public Property Get XMLDocumentInfo(diDesiredInfo As DocInfoConst) As Variant
Dim strDesiredInfo As String
Select Case diDesiredInfo
    Case diVERSION
        strDesiredInfo = "version"
    Case diENCODING
        strDesiredInfo = "encoding"
    Case diSTANDALONE
        strDesiredInfo = "standalone"
End Select

Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)

XMLDocumentInfo = xNode.Text
Set xNode = Nothing
End Property
'---------------------------------------------------
' Function
'   OpenXML
' Parameters
'   strSource As String : The complete path and
'       file name of the xml file to open
'   Type As tOpenXML : Supplies to the function
'       wether it should load from a file or a string
' What it does
'   Opens strFileName as an xml document ready for
'   input
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function OpenXML(strSource As String, Optional oType As OpenXMLDoc = oxFile) As Boolean
Set xDoc = New MSXML.DOMDocument

Select Case oType
    Case oxFile
        OpenXML = xDoc.Load(strSource)
    Case oxString
        OpenXML = xDoc.loadXML(strSource)
End Select
End Function
'---------------------------------------------------
' Function
'   ReadNode
' Parameters
'   strQuery As String : The query string to be
'       executed on the xml file
' What it does
'   Executes the query strQuery on the xml file
'   that has been opened and returns the value found
'   in the node
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function ReadNode(strQuery As String) As String
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

ReadNode = xNode.Text

Set xNode = Nothing

Exit Function
ErrHandle:
ReadNode = Null
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   ReadNodeXML
' Parameters
'   strQuery As String : The query string to be
'       executed on the xml file
' What it does
'   Executes the query strQuery on the xml file
'   that has been opened and returns the full xml
'   code for the node
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function ReadNodeXML(strQuery As String) As String
Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

ReadNodeXML = xNode.XML

Set xNode = Nothing
End Function
'---------------------------------------------------
' Public Property Get
'   XML
' Parameters
'   (None)
' What it does
'   Reads all of the XML data from the xml file and
'   outputs as a string
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Property Get XML() As String
XML = xDoc.XML
End Property
'---------------------------------------------------
' Function
'   WriteNode
' Parameters
'   strQuery As String : The query used to access the
'       node
'   Value As Variant : The value to be written to
'       the node
' What it does
'   Writes Value to the node specified by strQuery
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function WriteNode(strQuery As String, Value As Variant)
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

xNode.Text = Value

Set xNode = Nothing

Exit Function
ErrHandle:
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   ReadAttribute
' Parameters
'   strQuery As String : The query used to access the
'       node
'   strName As String : The name of the attribute to
'       read
' What it does
'   Returns the value of the attribute specified by
'   strName
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function ReadAttribute(strQuery As String, strName As String) As String
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMElement
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

ReadAttribute = xNode.getAttribute(strName)

Set xNode = Nothing

Exit Function
ErrHandle:
ReadAttribute = Null
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   WriteAttribute
' Parameters
'   strQuery As String : The query used to access the
'       node
'   strName As String : The name of the attribute
'       to be written
'   Value As Variant : The value to be written to
'       the node attribute
' What it does
'   Writes Value to the node attribute specified by
'   strName in a node specified by strQuery
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function WriteAttribute(strQuery As String, strName As String, Value As Variant)
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMElement
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

xNode.setAttribute strName, Value

Set xNode = Nothing

Exit Function
ErrHandle:
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   NodeCount
' Parameters
'   strQuery As String : The query used to access the
'       nodes
' What it does
'   Returns the count of the nodes
'   (not sure about functionality, no errors though)
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function NodeCount(strQuery As String) As Long
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMNodeList
Set xNode = xDoc.documentElement.selectNodes(strQuery)

NodeCount = xNode.length

Set xNode = Nothing

Exit Function
ErrHandle:
NodeCount = -1
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   Save
' Parameters
'   strFileName As String : Full path to the save
'       destination
' What it does
'   Saves the XML file to strFileName
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function Save(strFileName As String)
xDoc.Save strFileName
End Function
'---------------------------------------------------
' Function
'   DeleteNode
' Parameters
'   strQuery As String : Query to find the node
' What it does
'   Deletes a child node specified by strQuery
' Author
'   Anthony (11.18.02)
' Modified
'   (11.18.02)
'---------------------------------------------------
Public Function DeleteNode(strQuery As String)
On Error GoTo ErrHandle:

Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

xNode.parentNode.removeChild xNode

Set xNode = Nothing

Exit Function
ErrHandle:
Set xNode = Nothing
End Function
'---------------------------------------------------
' Function
'   MakeNode
' Parameters
'   strQuery As String : Query to find the parent
'       node
'   strName As String : Name of the new node
'   Optional Value As Variant : Value of the node
' What it does
'   Creates a new node in the XML file
' Author
'   Anthony (11.18.02)
' Modified
'   (11.19.02)
'---------------------------------------------------
Public Function MakeNode(strQuery As String, strName As String, Optional Value As Variant, Optional strAttributeName1 As String, Optional AttributeValue1 As Variant, Optional strAttributeName2 As String, Optional AttributeValue2 As Variant)
Dim xPNode As MSXML.IXMLDOMNode
Dim xCNode As MSXML.IXMLDOMElement
Set xPNode = xDoc.documentElement.selectSingleNode(strQuery)
Set xCNode = xDoc.createElement(strName)

If Not IsMissing(Value) Then
    xCNode.Text = Value
End If

xPNode.appendChild xCNode

If Not IsMissing(strAttributeName1) Then
    WriteAttribute strQuery & "/" & xCNode.nodeName, strAttributeName1, AttributeValue1
End If

If Not IsMissing(strAttributeName2) Then
    WriteAttribute strQuery & "/" & xCNode.nodeName, strAttributeName2, AttributeValue2
End If
End Function
'---------------------------------------------------
' Function
'   DeleteAttribute
' Parameters
'   strQuery As String : Query to find the parent
'       node
'   strName As String : Name of the attribute to be
'       deleted
' What it does
'   Deletes strAttributeName from the node specified
'   by strQuery
' Author
'   Anthony (11.19.02)
' Modified
'   (11.19.02)
'---------------------------------------------------
Public Function DeleteAttribute(strQuery As String, strAttributeName As String)
Dim xNode As MSXML.IXMLDOMElement
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

xNode.Attributes.removeNamedItem strAttributeName

Set xNode = Nothing
End Function

'---------------------------------------------------
' Function
'   GetColorValue
' Parameters
'   strQuery : Query used to locate the (x, x, x) node
' What it does
'   Parses a node found in the format (x, x, x) to be
'   used as an rgb value
'   Set the object to be colored equal to this
'   EX: frmMain.BackColor = XML.GetColorValue("query")
' Author
'   Anthony (11.20.02)
' Modified
'   (11.20.02)
'---------------------------------------------------
Public Function GetColorValue(strQuery As String) As Long
Dim xNode As MSXML.IXMLDOMNode
Set xNode = xDoc.documentElement.selectSingleNode(strQuery)

Dim Red As Integer, Green As Integer, Blue As Integer
Dim Data() As String

Data = Split(xNode.Text, ", ")

Red = Val(Right(Data(0), Len(Data(0)) - 1))
Green = Val(Data(1))
Blue = Val(Left(Data(2), Len(Data(2)) - 1))

GetColorValue = RGB(Red, Green, Blue)
End Function
'---------------------------------------------------
' Function
'   SetColorValue
' Parameters
'   strQuery : Query used to locate the (x, x, x) node
'   Red : Red color value
'   Green : Green color value
'   Blue : Blue color value
' What it does
'   Writes the color data with format (Red, Green, Blue)
'   To a specified node, allowing for easy skinning
' Author
'   Anthony (11.20.02)
' Modified
'   (11.20.02)
'---------------------------------------------------
Public Function SetColorValue(strQuery As String, Red As Integer, Green As Integer, Blue As Integer)
Dim strData As String
strData = "(" & Red & ", " & Green & ", " & Blue & ")"

WriteNode strQuery, Data
End Function
