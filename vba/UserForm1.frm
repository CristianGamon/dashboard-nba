VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Introduzca Cuotas"
   ClientHeight    =   2880
   ClientLeft      =   -90
   ClientTop       =   -330
   ClientWidth     =   3615
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButton1_Click()

Dim TextBoxVisitor As Range
Dim TextBoxLocal As Range

Set TextBoxVisitor = Hoja3.Range("f376").End(xlDown).Offset(1, -3)
Set TextBoxLocal = Hoja3.Range("f376").End(xlDown).Offset(1, -1)

TextBox4.Value = TextBoxVisitor
TextBox2.Value = TextBoxLocal

End Sub


Private Sub CommandButton2_Click()

Dim CuotaLocal As Variant
Dim CuotaVisitor As Variant

CuotaLocal = TextBox6.Value
CuotaVisitor = TextBox5.Value

Hoja3.Range("f376").End(xlDown).Offset(1, 0) = CuotaVisitor
Hoja3.Range("f376").End(xlDown).Offset(0, 1) = CuotaLocal
TextBox6.Value = ""
TextBox5.Value = ""
TextBox4.Value = ""
TextBox2.Value = ""

End Sub

Private Sub TextBox5_Change()

    Dim CuotaVisitor As Variant
    Dim Largo As Integer
    Dim i As Integer
    Dim Caracter As Variant
    
    CuotaVisitor = Me.TextBox5.Value
    Largo = VBA.Len(CuotaVisitor)
    
    For i = 1 To Largo
        Caracter = Mid(VBA.CStr(CuotaVisitor), i, 1)
        
        If Caracter <> "" Then
            If Caracter < VBA.Chr(46) Or Caracter > VBA.Chr(57) Then
                CuotaVisitor = VBA.Replace(CuotaVisitor, Caracter, "")
                Me.TextBox5.Value = CuotaVisitor
            Else
            End If
        End If
    Next i
        
End Sub

Private Sub TextBox6_Change()

    Dim CuotaLocal As Variant
    Dim Largo As Integer
    Dim i As Integer
    Dim Caracter As Variant
    
    CuotaLocal = Me.TextBox6.Value
    Largo = VBA.Len(CuotaLocal)
    
    For i = 1 To Largo
        Caracter = Mid(VBA.CStr(CuotaLocal), i, 1)
        
        If Caracter <> "" Then
            If Caracter < VBA.Chr(46) Or Caracter > VBA.Chr(57) Then
                CuotaLocal = VBA.Replace(CuotaLocal, Caracter, "")
                Me.TextBox6.Value = CuotaLocal
            Else
            End If
        End If
    Next i
End Sub

Private Sub UserForm_Click()

End Sub
