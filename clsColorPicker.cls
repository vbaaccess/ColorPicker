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
'----------------------------------------------------------------------------------------------------------------------------------

Private Const CurrentModName = "clsColorPicker"
Private Const INIT_STRING_EVENT = "[Event Procedure]"

Private objHandle

Public Event PrzypiszUchwytDoObiektu(uchwytDoObiektu As Long)   'to myself

Private WithEvents ctrEtyColor As Label

Public Sub InitializeToForm(ByRef setObjForm As Form)
    Dim l_objHandle As Long
    
    If objHandle = 0 Then
        RaiseEvent PrzypiszUchwytDoObiektu(l_objHandle)
        objHandle = CLng(l_objHandle)
    End If
    
    Call AttachTheEventToLabels(setObjForm)
    
End Sub

Private Sub AttachTheEventToLabels(objForm As Form)
    Debug.Print "UNDER CONSTRUCTION"
End Sub
