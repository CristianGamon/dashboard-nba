VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Seleccione Fecha"
   ClientHeight    =   1770
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   3465
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim FechaCompleta As Long
Dim Combo1, combo2, combo3

Combo1 = Me.ComboBox1.Value
combo2 = Me.ComboBox2.Value
combo3 = Me.ComboBox3.Value

If Combo1 = "" Then
    'Me.ComboBox1.BackColor = RGB(295, 110, 110)
    MsgBox "Debe rellenar todos los campos", vbExclamation
ElseIf combo2 = "" Then
    'Me.ComboBox2.BackColor = RGB(295, 110, 110)
    MsgBox "Debe rellenar todos los campos", vbExclamation
ElseIf combo3 = "" Then
    'Me.ComboBox3.BackColor = RGB(295, 110, 110)
    MsgBox "Debe rellenar todos los campos", vbExclamation
Else

FechaCompleta = Me.ComboBox3.Value & Me.ComboBox2.Value & Me.ComboBox1.Value


DescargaResultados FechaCompleta

End If

End Sub

Private Sub UserForm_Initialize()

With Me
    .ComboBox1.RowSource = "TablaDias[Día]"
    .ComboBox2.RowSource = "TablaMeses[Mes]"
    .ComboBox3.RowSource = "TablaAños[Año]"
End With

End Sub
