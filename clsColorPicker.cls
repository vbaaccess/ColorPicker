Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : clsColorPicker
' Author    : paweld
' Date      : 2019-10-01
' Purpose   : Universal color selection form
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2019-10-01        , PD : Creating the first version of the class
' 2019-11-20        , PD : Beta version - stable
'----------------------------------------------------------------------------------------------------------------------------------

Private Const CurrentModName = "clsColorPicker"
Private Const INIT_STRING_EVENT = "[Event Procedure]"
Private Const ColorPickerFormName = "frmColorPicker"

Private objHandle

Public Event AssignAHandleToTheObject(ObjectHolder As Long)   'to myself
Public Event ColorUpdated(newColor As Long)
 
Private arrColors() As Long
Dim pomArray(505, 3) As Long

Private m_DialogWindowMode As Boolean
Private m_SelectedColor As Long
Private WithEvents ctrEtyColor As Label

Private WithEvents ctrPreviewLabel1 As Label
Private WithEvents ctrPreviewLabel2 As Label

Private Sub Class_Initialize()
    '--- constructor ---
    m_DialogWindowMode = True
End Sub

Public Property Get DialogWindowMode() As Boolean
    DialogWindowMode = m_DialogWindowMode
End Property

Public Property Let DialogWindowMode(bDialogWindowMode As Boolean)
    m_DialogWindowMode = bDialogWindowMode
End Property

Public Property Get LastSelectedColor() As Long
    LastSelectedColor = m_SelectedColor
End Property

Public Property Let LastSelectedColor(newColor As Long)
    If m_SelectedColor <> newColor Then
        m_SelectedColor = newColor
        RaiseEvent ColorUpdated(newColor)
    End If
End Property

Public Function etyColorOnDblClick(ByRef objForm As Form, sCtrName As String)
    Dim ctrBackColor As Long
    
    ctrBackColor = objForm.Controls(sCtrName).BackColor
    
    If ctrBackColor > 0 Then
        Call SetPreviewSelectedColor(ctrBackColor)
    End If
End Function

Private Sub SetPreviewSelectedColor(newBackColor As Long)
    If ctrPreviewLabel1 Is Nothing Or ctrPreviewLabel2 Is Nothing Then
        'blad !
    Else
        ctrPreviewLabel1.BackColor = newBackColor
        ctrPreviewLabel2.BackColor = newBackColor
        ctrPreviewLabel1.BackStyle = 1
        ctrPreviewLabel2.BackStyle = 1
    End If
End Sub

Public Sub INIT()
    Call InitHandle
End Sub

Private Sub InitHandle()
    Dim l_objHandle As Long
    
    If objHandle = 0 Then
        RaiseEvent AssignAHandleToTheObject(l_objHandle)
        objHandle = CLng(l_objHandle)
    End If
End Sub

Public Sub ShowColorPickerForm()
    Dim WindowMode As AcWindowMode
    
    Call InitHandle
                                   
    If m_DialogWindowMode Then
        WindowMode = acDialog
    Else
        WindowMode = acWindowNormal  ' Default
    End If
    
   DoCmd.OpenForm ColorPickerFormName, , , , , WindowMode, objHandle
End Sub

Public Sub INIT_frmColorPicker(ByRef setObjForm As Form)
    Call InitHandle
    
    '--- 1 ---
    Call UtworzTabliceKolorow(pomArray)
   
   '--- attach the event to controls ---
    If Not ctrPreviewLabel1 Is Nothing Then Set ctrPreviewLabel1 = Nothing
    Set ctrPreviewLabel1 = setObjForm.Controls("etyPreviewLabel1")
    ctrPreviewLabel1.OnClick = INIT_STRING_EVENT
    ctrPreviewLabel1.BackStyle = 0
    
    If Not ctrPreviewLabel2 Is Nothing Then Set ctrPreviewLabel2 = Nothing
    Set ctrPreviewLabel2 = setObjForm.Controls("etyPreviewLabel2")
    ctrPreviewLabel2.OnClick = INIT_STRING_EVENT
    ctrPreviewLabel2.BackStyle = 0
    
    Call AttachTheEventToLabels(setObjForm)
    
End Sub

Private Sub AttachTheEventToLabels(objForm As Form)
    Dim sCtrName As String
    ReDim arrColors(1 To 9, 1 To 56)
    Dim LpRect&, LpEty
    Dim iLpColor As Long
    Dim iColor As Long
    
    Dim sEventFunction As String
    
    For LpRect = 1 To 9
        For LpEty = 1 To 56
            
            sCtrName = "etyC" & LpRect & Right("0" & LpEty, 2)
            iLpColor = iLpColor + 1
            If iLpColor <= 505 Then
                iColor = pomArray(iLpColor, 3)
            End If
            
            objForm.Controls(sCtrName).Caption = vbNullString
            objForm.Controls(sCtrName).BackColor = iColor
            
            sEventFunction = "=etyC_DblClick(" & Chr(34) & sCtrName & Chr(34) & ")"
            objForm.Controls(sCtrName).OnDblClick = sEventFunction
            objForm.Controls(sCtrName).OnClick = sEventFunction
            
        Next LpEty
    Next LpRect

End Sub

Private Sub ctrPreviewLabel1_Click()
    Call ctrPreviewLabelClick
End Sub

Private Sub ctrPreviewLabel2_Click()
    Call ctrPreviewLabelClick
End Sub

Private Sub ctrPreviewLabelClick()
    Dim sPrompt$, sTitle$
    Dim MsgBoxResult As VbMsgBoxResult
    
    sPrompt = "Do you want to save the selected color?"
    sTitle = "ATTENTION"
    
    MsgBoxResult = VBA.MsgBox(sPrompt, vbQuestion + vbYesNo, sTitle)
    
    If MsgBoxResult = vbYes Then
        Me.LastSelectedColor = ctrPreviewLabel1.BackColor
    End If
End Sub
