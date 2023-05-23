VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_InicioSesion 
   Caption         =   "Inicio de sesión"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   OleObjectBlob   =   "frm_InicioSesion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_InicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objDatePicker As Object

Private Sub cmd_downloadView_Click()
    Rem validadcion
    
    If txt_username = Empty Or txt_password = Empty Or cmb_listView.ListIndex < 0 Then
        MsgBox "Debe completar los campos y seleccionar un view", vbInformation, "Inicio de sesión"
        txt_username.SetFocus
        Exit Sub
    End If
 
    user = txt_username.Text
    password = txt_password.Text
    
    With cmb_listView
        view = .List(.ListIndex, 1)
        viewText = .List(.ListIndex, 0)
        
        If Not objDatePicker Is Nothing Then
            date_view = objDatePicker.Value
        End If
        
        Debug.Print view
    End With
    
    Debug.Print user
    Debug.Print password
    Debug.Print cmb_listView.List(cmb_listView.ListIndex, 0)
    Debug.Print cmb_listView.List(cmb_listView.ListIndex, 1)
'    Debug.Print DTPicker1.Value
    
    Unload Me
'    Call query_view1
    Call main.query_view2
   
End Sub

Private Sub cmd_start_Click()

    If txt_username = Empty Or txt_password = Empty Then
        MsgBox "Debe completar los campos", vbInformation, "Inicio de sesión"
        txt_username.SetFocus
        Exit Sub
    End If
 
    user = txt_username.Text
    password = txt_password.Text
    
    Unload Me
    
    Call main.view_client
   
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    
    Rem Instancias objeto de clase dataPicker
    Set objDatePicker = CreateObject("MSComCt12.DTPicker")
    
    cmd_exit.Cancel = True
    
    txt_username.TabIndex = 0
        
    With txt_password
        .TabIndex = 1
        .PasswordChar = "*"
    End With

    
    If calling = 1 Then
        frm_InicioSesion.Height = 264
        Frame1.Height = 234
        
        With cmb_listView
            .TabIndex = 2
            .RowSource = "lista_valores"
        End With
        
        Rem usar el control DataPicker
        If Not objDatePicker Is Nothing Then
             With objDatePicker
                .TabIndex = 3
                .Top = 163
            End With
'        DTPicker1.TabIndex = 3
        End If
        
        With cmd_downloadView
            .TabIndex = 4
            .Default = True
        End With
        
        cmd_start.Enabled = False

    ElseIf calling = 2 Then
        frm_InicioSesion.Height = 200
        Frame1.Height = 160
        
        Label3.Visible = False
        cmb_listView.Visible = False
        cmd_downloadView.Visible = False
'        DTPicker1.Enabled = False
        
        With cmd_start
            .Top = 120
            .Default = True
            .TabIndex = 2
        End With
        
        txt_username = "dannin"
        txt_password = "Procesos2023*"
        
    End If

End Sub

