VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   2565
      TabIndex        =   0
      Top             =   765
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim oDom As New MSXML2.DOMDocument30
   Dim oElement As MSXML2.IXMLDOMElement
   Dim oNode As MSXML2.IXMLDOMNode
   Dim oNodeList As MSXML2.IXMLDOMNodeList
   With oDom
       Set oNode = .appendChild(.createElement("MyElementRoot"))
       Set oNode = oNode.appendChild(.createElement("MyElement")) ' This would be the field name
       oNode.Text = "This is the element value" ' You could set this equal to your form value
   End With
   oDom.save "c:\myxml.xml" 'If you want to persist it
   Set oNodeList = oDom.selectNodes("//MyElement") 'this will return a list of MyElement Elements
                                                                                    ' if you 'had more than 1
   Set oNode = oDom.selectSingleNode("/MyElementRoot/MyElement") 'This would be the first one or you
                                                                                                                'could loop through the nodelist
   
  Set oDom = Nothing
  Set oNodeList = Nothing
  Set oNode = Nothing
End Sub
