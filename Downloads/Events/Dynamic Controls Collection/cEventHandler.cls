VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Public Event Click(ByVal Index As Integer, ByVal ObjectType As String)
Public Event Changed(ByVal Index As Integer, ByVal ObjectType As String)

Public mInstance As Integer

Public Sub ObjectClicked(ByVal Index As Integer, ByVal ObjectType As String)
     
    RaiseEvent Click(Index, ObjectType)

End Sub

Public Sub ObjectChanged(ByVal Index As Integer, ByVal ObjectType As String)
     
    RaiseEvent Changed(Index, ObjectType)

End Sub

Private Sub Class_Initialize()

    mInstance = mInstance + 1
    
End Sub
