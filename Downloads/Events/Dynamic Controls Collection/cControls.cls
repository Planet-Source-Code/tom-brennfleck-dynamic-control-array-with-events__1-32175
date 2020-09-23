VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Private WithEvents oLabel As Label
Attribute oLabel.VB_VarHelpID = -1
Private WithEvents oCheck As CheckBox
Attribute oCheck.VB_VarHelpID = -1
Private WithEvents oText As TextBox
Attribute oText.VB_VarHelpID = -1

Private m_CallingForm As Form
Private m_Container As Object

Public Property Set oContainer(cnt As Object)

    Set m_Container = cnt
    
End Property

Public Property Set oForm(F As Form)

    Set m_CallingForm = F
    
End Property

Public Sub Draw(ByVal ObjectType As String, _
                  ByVal Name$, _
                  ByVal Top%, _
                  ByVal Left%, _
                  ByVal Height%, _
                  ByVal Width%)
'***************************************************************************
'*  Name         : Draw
'*  Description  : Draw the object on the screen
'*  Parameters   : ObjectType As String, Name$,
'*               : Top%, Left%, Height%, Width%
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************

        Select Case UCase(ObjectType)
            Case "LABEL"
                Set oLabel = m_CallingForm.Controls.Add("VB.Label", Name, m_Container)
                Set m_CallingForm = Nothing          'Important Line of Code
                Set m_Container = Nothing
                'Set oLabel = Nothing
                ' we set the properties
                oLabel.AutoSize = False
                oLabel.BackStyle = 1
                oLabel.Caption = "Label Box"
                SetPosition oLabel, Top, Left, Height, Width
            Case "CHECKBOX"
                Set oCheck = m_CallingForm.Controls.Add("VB.CheckBox", Name, m_Container)
                Set m_CallingForm = Nothing          'Important Line of Code
                Set m_Container = Nothing
                'Set oCheck = Nothing
                oCheck.Caption = "Check Box"
                oCheck.Value = 0
                SetPosition oCheck, Top, Left, Height, Width
            Case "TEXTBOX"
                Set oText = m_CallingForm.Controls.Add("VB.TextBox", Name, m_Container)
                Set m_CallingForm = Nothing          'Important Line of Code
                Set m_Container = Nothing
                'Set oText = Nothing
                oText.Text = ""
                SetPosition oText, Top, Left, Height, Width
            Case Else
        End Select
        
                
End Sub

Private Sub SetPosition(oObject As Object, _
                  ByVal Top%, _
                  ByVal Left%, _
                  ByVal Height%, _
                  ByVal Width%)
'***************************************************************************
'*  Name         : SetPosition
'*  Description  : set the position of the object
'*  Parameters   : oObject As Object, Top%,
'*               : Left%, Height%, Width%
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************

        'then the position
        oObject.Top = Top
        oObject.Left = Left
        oObject.Height = Height
        oObject.Width = Width

        'and make visible
        oObject.Visible = True

End Sub

Private Sub oLabel_Click()
'***************************************************************************
'*  Name         : oLabel_Click
'*  Description  : Label Click event
'*  Parameters   : None
'*  Returns      : Generates an event in the event handler
'*  Called From  :
'*  Author       :  T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************
Dim Index As Integer
        
        Index = ObjectChange(oLabel)
        
        ' raises the event
        EventHandler.ObjectClicked Index, "Label"
        
End Sub

Private Sub oText_Change()
'***************************************************************************
'*  Name         : oText_Change
'*  Description  :
'*  Parameters   : None
'*  Returns      : Generates an event in the event handler
'*  Called From  :
'*  Author       :  T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************
Dim Index As Integer
        
        Index = ObjectChange(oText)
        
        ' raises the event
        EventHandler.ObjectChanged Index, "TextBox"
        
End Sub

Private Sub oCheck_Click()
'***************************************************************************
'*  Name         : oCheck_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Generates an event in the event handler
'*  Called From  :
'*  Author       :  T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************
Dim Index As Integer
        
        Index = ObjectChange(oCheck)
        
        ' raises the event
        EventHandler.ObjectClicked Index, "CheckBox"
        
End Sub

Public Sub SetObjectProperty(oObject As Object, Property As String, Optional Value As Variant)

    Select Case Property
        Case "BackColor"
            If IsMissing(Value) Then
                Value = &H80000005
            End If
            oObject.BackColor = Value
        Case "Text"
            If IsMissing(Value) Then
                Value = ""
            End If
            oObject.Text = Value
        Case Else
    End Select
    
End Sub

Private Function ObjectChange(oObject As Object) As Integer
'***************************************************************************
'*  Name         : ObjectChange
'*  Description  : Return the index of the object that has chnaged
'*  Parameters   : oObject As Object
'*  Returns      : Integer = xx --> index of the object
'*               : Integer = -1 --> no object found
'*  Called From  :
'*  Author       :  T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************
Dim Index As Integer
Dim sTmp As Variant

        sTmp = Split(oObject.Name, "_")
        
        If UBound(sTmp) > 0 Then
            Index = sTmp(1)
        Else
            Index = -1
        End If

        ObjectChange = Index
        
        
End Function

Public Function ClearObject(F As Form, oCollection As Collection, _
                        Optional Name As String = "", _
                        Optional Index As Integer = -1) As Boolean
'***************************************************************************
'*  Name         : ClearObject
'*  Description  :
'*  Parameters   : F As Form, oCollection As Collection,
'*               : Optional Name As String = ""
'*               : Optional Index As Integer = -1
'*  Returns      : True on Success
'*               : False on Failure
'*               :
'*  Assumes      : if Name = "" And Index = -1 --> delete all controls auto generated
'*               : if Name = "" And Index = SomeIndex --> Do Nothing
'*               : if Name = "SomeName" And Index = -1 --> delete all controls of name
'*               : if Name = "SomeName" And Index = SomeIndex --> delete control of name and Index
'*               :
'*  Called From  :
'*  Author       :  T. Brennfleck
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         : Add error handling, as this procedure always returns true
'*
'***************************************************************************
On Error GoTo ClearObjectError
Dim bSucces As Boolean

    bSucces = False
    
    If Name = "" Then
        If Index = -1 Then
            DeleteControls F, oCollection
            bSucces = True
        End If
    Else
        If Index = -1 Then
            DeleteControls F, oCollection, Name
        Else
            DeleteControls F, oCollection, Name & "_" & Index, Index
        End If
    End If
    
    
ClearObjectExit:
    Exit Function
    
ClearObjectError:
    Resume ClearObjectExit
    
End Function

Private Sub DeleteControls(F As Form, oCollection As Collection, _
                            Optional Name As String = "", _
                            Optional Index As Integer)
'***************************************************************************
'*  Name         : DeleteControls
'*  Description  : Delete a control or controls from the form and the collection
'*  Parameters   : F As Form, oCollection As Collection,
'*               : Optional Name As String = "", Optional Index As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Administrator
'*  Date         : 28 Feb 2002
'*               -------------------------------------------------------------
'*  History      :
'*   28 Feb 2002 : Initial Version
'*
'*  ToDo         :
'*
'***************************************************************************
Dim C As Control
Dim ControlName As String
Dim ControlIndex As String
Dim ControlArray As Variant

    ControlArray = Split(Name, "_")
    Select Case UBound(ControlArray)
        Case -1
            For Each C In F.Controls
                'remove all controls
                If InStr(1, C.Name, "_") > 0 Then
                    oCollection.Remove C.Name
                    F.Controls.Remove C.Name
                End If
            Next
        Case 0
            For Each C In F.Controls
                'remove a selected set of controls
                ControlName = Split(C.Name, "_")(0)
                If ControlName = Name Then
                    oCollection.Remove C.Name
                    F.Controls.Remove C.Name
                End If
            Next
        Case 1
            For Each C In F.Controls
                'remove a control
                ControlName = C.Name
                If ControlName = Name Then
                    oCollection.Remove C.Name
                    F.Controls.Remove C.Name
                End If
            Next
        Case Else
    End Select
    
End Sub
















