VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const CurrentFormName = "frmColorPicker"

Private arrColors() As Long

Dim arrColors(505, 3) As Long

Private Sub Form_Open(Cancel As Integer)
    '-- 1 --
    Call UtworzTabliceKolorow(arrColors)
End Sub

Private Sub Form_Load()
    '-- 2 --
    Call InitColorsArray
End Sub

Private Sub UstawWskaznik(obj As Object)
    If IsNumeric(obj.Tag) Then
        If CLng(obj.Tag) > 0 Then
            If Screen.MousePointer = 0 Then
                Call MouseCursor(IDC_HAND)
            End If
        End If
    End If
End Sub

Private Sub CmdWyjscie_Click()
    DoCmd.Close acForm, CurrentFormName
End Sub

Private Sub InitColorsArray()
    Dim sCtrName As String
    ReDim arrColors(1 To 9, 1 To 56)
    Dim LpRect&, LpEty
    Dim iLpColor As Long
    Dim iColor As Long
    
    For LpRect = 1 To 9
        For LpEty = 1 To 56
            
            sCtrName = "etyC" & LpRect & Right("0" & LpEty, 2)
            iLpColor = iLpColor + 1
            If iLpColor <= 505 Then
                iColor = arrColors(iLpColor, 3)
            End If
            
            Me.Controls(sCtrName).Caption = vbNullString
            Me.Controls(sCtrName).BackColor = iColor
            
        Next LpEty
    Next LpRect

End Sub

Sub OdswiezDane()

    Call UstawStartoweKoloryPrzyciskow
    Call OdswiezOpisyPrzyciskow
    
    Call AktualizujDane
    Call OdswiezOpisyPrzyciskow

End Sub


