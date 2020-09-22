VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Example program for vbXML"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReloadXML 
      Caption         =   "Reload the skin.xml file"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Make sure you use the Color Codes!!! =)"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "To change the color of the form just edit the <mainbg> node in the skin.xml file."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "The Color Code shown here is identical to the RGB function in VB. "
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblCCode 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Form Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Form Background Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XML As New clsXML

Dim Red As Integer, Green As Integer, Blue As Integer

Private Sub cmdReloadXML_Click()
GetSkin
End Sub

Private Sub Form_Load()
'---------------------------------------------------
' cXML uses the MSXML feature XPath.  Here is a quick
' example of how I would query a node (used in
' ReadNode, ReadNodeXML, and WriteNode):
' To access a node:
'   "/parent1/childofparent1/childofparent2"
' Here is a quick explination:
' I am using test.xml for this example.  For ease of
' this example I have pasted the contents here:
'
'  <test>
'     <text>
'         <hello>Hello</hello>
'         <bye>Goodbye</bye>
'     </text>
'  </test>
'
' To access the <hello> node, you would use this
' query:
'   "/test/text/hello"
'
' To access the <bye> node, you would use this query:
'   "/test/text/bye"
'
' In the last example (where we queried the <bye> node)
' "/test/" is parent1, "/text/" is childofparent1,
' and "/bye" is childofparent2
'
' Take note: you can have multiple child nodes
' (I dont know the exact count)
'---------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' This example creates a skinable form with vbXML
' (This is not a great skinable form, but it shows
' adequetly what vbXML can do)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'---------------------------------------------------
' See the GetSkin method to see how nodes are read
'---------------------------------------------------
GetSkin
End Sub

Public Sub GetSkin()
'---------------------------------------------------
' Open the skin file and get ready for action :)
'---------------------------------------------------
XML.OpenXML App.Path & "/skin.xml"

'---------------------------------------------------
' Change the forms background color to the one(s)
' specified in the skin.xml file
' Red, Green, and Blue are different nodes
' I should add a way to parse this to vbXML, maybe
' in the next release of the wrapper
'---------------------------------------------------

frmMain.BackColor = XML.GetColorValue("/skin/colors/mainbg")
lblCCode.Caption = XML.ReadNode("/skin/colors/mainbg")

'---------------------------------------------------
' Read in the forms caption
'---------------------------------------------------
frmMain.Caption = XML.ReadNode("/skin/text/main/caption")
txtCaption = frmMain.Caption
End Sub

Private Sub txtCaption_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

XML.WriteNode "/skin/text/main/caption", txtCaption
frmMain.Caption = XML.ReadNode("/skin/text/main/caption")
End Sub
