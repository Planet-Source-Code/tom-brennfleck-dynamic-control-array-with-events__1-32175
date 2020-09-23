VERSION 5.00
Begin VB.Form frm1 
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetColortBox 
      Caption         =   "Set Color Of Text Box"
      Height          =   315
      Left            =   3060
      TabIndex        =   8
      Top             =   2580
      Width           =   2895
   End
   Begin VB.CommandButton cmdSetTextBox 
      Caption         =   "Set Text Of Text Box"
      Height          =   315
      Left            =   3060
      TabIndex        =   7
      Top             =   2940
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "standard Text Box"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Standard Check Box"
      Height          =   255
      Left            =   1500
      TabIndex        =   5
      Top             =   2280
      Width           =   2085
   End
   Begin VB.CommandButton cmdDeleteCheck2 
      Caption         =   "Delete Only Check Box 2"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3300
      Width           =   2895
   End
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "Delete All Auto Generated Controls"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Width           =   2895
   End
   Begin VB.CommandButton cmdDeleteCheck 
      Caption         =   "Delete All Check Boxes"
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   3300
      Width           =   2895
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "[Me First] -- Auto Generate Controls"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2580
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Standard Label"
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   2280
      Width           =   1245
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Copyright T. Brennfleck SamSa Consulting Group Pty. Ltd. 2001
'   This code is copyrighted and has limited warranties.
'   Terms of Agreement:
'   By using this code, you agree to the following terms...
'   1) You may use this code in your own programs
'      (and may compile it into a program and distribute it in compiled
'      format for langauges that allow it) freely and with no charge.
'   2) You MAY NOT redistribute this code (for example to a web site, or code library, or book)
'      without written permission from the original author.
'      Failure to do so is a violation of copyright laws.
'   3) If you use this code for a commercial product I would like to have an email
'      stating the product that it is being used in.

Private WithEvents EH As cEventHandler
Attribute EH.VB_VarHelpID = -1

Private oObject As cControl
Private mColObjects As Collection

Private Sub CreateControls()
Dim Index As Integer
   
    Set mColObjects = New Collection
    
    For Index = 0 To 3
        Set oObject = New cControl
        Set oObject.oForm = Me
        Set oObject.oContainer = Me
        
        ' Add Item to Collection
        oObject.Draw "Label", "Label_" & Index, 0 + (500 * Index), 0, 250, 1000
        mColObjects.Add oObject, "Label_" & Index
        Set oObject = Nothing
    Next
    
    For Index = 0 To 3
        Set oObject = New cControl
        Set oObject.oForm = Me
        Set oObject.oContainer = Me
        
        ' Add Item to Collection
        oObject.Draw "CheckBox", "Check_" & Index, 0 + (500 * Index), 1500, 250, 2000
        mColObjects.Add oObject, "Check_" & Index
        Set oObject = Nothing
    Next
    
    For Index = 0 To 3
        Set oObject = New cControl
        Set oObject.oForm = Me
        Set oObject.oContainer = Me
        
        ' Add Item to Collection
        oObject.Draw "TextBox", "Text_" & Index, 0 + (500 * Index), 4000, 250, 2000
        mColObjects.Add oObject, "Text_" & Index
        Set oObject = Nothing
    Next
    
End Sub


Private Sub EventHandler_Changed(ByVal Index As Integer, ByVal ObjectType As String)

    MsgBox "Event Generated" & vbCrLf & vbCrLf & "Index: " & Index & vbCrLf & "Object Type: " & ObjectType
    
End Sub

Private Sub EventHandler_Click(ByVal Index As Integer, ByVal ObjectType As String)

    MsgBox "Event Generated" & vbCrLf & vbCrLf & "Index: " & Index & vbCrLf & "Object Type: " & ObjectType
    
End Sub

Private Sub cmdCreate_Click()
'create the test controls

    CreateControls
    
End Sub

Private Sub cmdDeleteAll_Click()
'delete all test controls

    Set oObject = New cControl
    oObject.ClearObject Me, mColObjects
    Set oObject = Nothing
    
End Sub

Private Sub cmdDeleteCheck_Click()
'clear the check boxes only

    Set oObject = New cControl
    oObject.ClearObject Me, mColObjects, "Check"
    Set oObject = Nothing
    
End Sub

Private Sub cmdDeleteCheck2_Click()
'delete check box 2
    
    Set oObject = New cControl
    oObject.ClearObject Me, mColObjects, "Check", 2
    Set oObject = Nothing
    
End Sub


Private Sub cmdSetColortBox_Click()

    Set oObject = New cControl
    If Me.Controls("text_2").BackColor = &HFF& Then
        oObject.SetObjectProperty Me.Controls("text_2"), "BackColor"
    Else
        oObject.SetObjectProperty Me.Controls("text_2"), "BackColor", &HFF&
    End If
    Set oObject = Nothing

End Sub

Private Sub cmdSetTextBox_Click()
    
    Set oObject = New cControl
    If Me.Controls("text_1").Text = "" Then
        oObject.SetObjectProperty Me.Controls("text_1"), "Text", "New Text"
    Else
        oObject.SetObjectProperty Me.Controls("text_1"), "Text"
    End If
    Set oObject = Nothing

End Sub

Private Sub EH_Changed(ByVal Index As Integer, ByVal ObjectType As String)

    MsgBox "Event Generated" & vbCrLf & vbCrLf & "Index: " & Index & vbCrLf & "Object Type: " & ObjectType
    
End Sub

Private Sub EH_Click(ByVal Index As Integer, ByVal ObjectType As String)

    MsgBox "Event Generated" & vbCrLf & vbCrLf & "Index: " & Index & vbCrLf & "Object Type: " & ObjectType
    
End Sub

Private Sub Form_Load()

    Set EH = EventHandler
    
End Sub
