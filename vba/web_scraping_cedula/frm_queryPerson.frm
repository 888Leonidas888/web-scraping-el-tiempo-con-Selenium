VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_queryPerson 
   Caption         =   "Consulta por cédula"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "frm_queryPerson.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_queryPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const APP_NAME = "Consultar por cédula"

Private Sub cmd_search_Click()
    If IsNumeric(txt_cedula) Then
'        MsgBox "Aun falta implementar esta modulo", vbCritical, APP_NAME
        With lbl_fullName
            .Caption = vbEmpty
            .Caption = searchPersonForCedula(txt_cedula)
            .ForeColor = vbBlue
        End With
    Else
        MsgBox "Se espera datos númericos", vbCritical, APP_NAME
    End If
End Sub

Private Sub cmd_searchBatch_Click()
    
    Const NAME_SH = "lista_cedulas"
    Dim rng As Range
    Dim c As Integer, i As Integer
    Dim cedula As Variant

    If ActiveSheet.Name <> NAME_SH Then
        MsgBox "La hoja actual no es " & NAME_SH, vbCritical, APP_NAME
        Exit Sub
    End If
    
    With WorksheetFunction
        c = .CountA(Worksheets(NAME_SH).Range("A:A"))
    End With
    
    For i = 2 To c
        cedula = Range("A" & i).Value
        
        If IsNumeric(cedula) Then
            With WorksheetFunction
                Range("B" & i).Value = .Proper(searchPersonForCedula(cedula))
                Range("B" & i).Font.Color = vbBlack
            End With
        Else
            With Range("B" & i)
                .Value = "Cédula no valida"
                .Font.Color = vbRed
            End With
        End If
    Next i
    
    MsgBox "Consulta masiva terminada", vbInformation, APP_NAME
End Sub


Private Sub UserForm_Initialize()
    txt_cedula = 1052406961
    
'    Me.BackColor = RGB(166, 199, 248)
'    For i = 1 To 3
'        Me.Controls("Frame" & i).BackColor = RGB(91, 153, 243)
'    Next i
End Sub
