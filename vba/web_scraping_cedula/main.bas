Attribute VB_Name = "main"
Public user As String
Public password As String
Public date_view As Date
Public view As Long
Public viewText As String
Public calling As Byte
Public Const NAME_HJ As String = "consolidado"

Sub show_form1(control As IRibbonControl)
    Rem descarga de vistas
    calling = 1
    frm_InicioSesion.Show
End Sub

Sub show_form2(control As IRibbonControl)
    calling = 2
    frm_InicioSesion.Show
End Sub
Sub show_form_cedula(control As IRibbonControl)
    MsgBox "Este porceso no esta habilitado para este libro", vbInformation
    Exit Sub
    frm_queryPerson.Show
End Sub

Sub query_view1()
    Rem rutina en desuso, utiliza la libreria Microsoft Internet Controls(Internet Explorer-descontinuado)
    Rem  Use rutina query_view1 basada en SeleniumBasic
    
    Dim IE As New SHDocVw.InternetExplorer
    Dim url As String, url2 As String

    Dim i As Integer
    
    On Error GoTo Cath
    
    url = "https://vencimientosceet.eltiempo.com.co/CEETVencimientos/Autenticacion/IniciarSesion.aspx"
    url2 = "https://vencimientosceet.eltiempo.com.co/CEETVencimientos/ConsultaVistas/Default.aspx"
    
    With IE
        .Navigate url
        .Visible = True
        
        Application.wait Now() + TimeValue("00:00:02")
   
        .Document.getelementbyId("content_content_nombreUsuarioTextBox").Value = user
        .Document.getelementbyId("content_content_claveTextBox").Value = password
        .Document.getelementbyId("content_content_iniciarSesionButton").Click


        Do While .LocationURL = url
            'espera
        Loop

'        .Document.getelementbyid("ui-accordion-1-panel-2").Click
        .Navigate2 url2
        
        Application.wait Now() + TimeValue("00:00:02")


    
        .Document.getelementbyId("content_content_vistaDropDownList").Value = view


        
        .Document.getelementbyId("content_content_generarButton").Click

        Application.wait Now() + TimeValue("00:00:03")
                                            
        .Document.getelementbyId("content_content_descargarButton").Click
        
        MsgBox "Se ha abierto una ventana de Internet explorer,presione guardar para continuar ", vbOKOnly + vbDefaultButton1 + vbInformation, "Ingrese a Interner Explorer"
        Exit Sub
    End With
Cath:
    MsgBox Err.Description & vbCrLf & "Por favor intenta de nuevo", vbCritical, Err.Number
    
End Sub
Sub query_view2()
    
    Rem DESCARGAR VISTAS
    Rem Descargar Archivos Excel usando SeleniumBasic
    Rem cada sierto posiblemente necesite actualizar el chromedriver.exe,dado que se actualiza el navegador también
    Rem la ruta posible donde se encuentra el chromedriver C:\Users\MI_USUARIO\AppData\Local\SeleniumBasic
    
    Rem este procedimiento esta asociado al botón de la ribbon "Descarga de vistas"
    
    Dim g As New WebDriver
    Dim url As String
    Dim listDrow As SelectElement
    
'    user = ""
'    password = ""
'    date_view = 15
      
    On Error GoTo Cath

    url = "https://vencimientosceet.eltiempo.com.co/CEETVencimientos/ConsultaVistas/Default.aspx"
    
    With g
        .Start "chrome"
        .Get url
        
        .FindElementById("content_content_nombreUsuarioTextBox").SendKeys user
        .FindElementById("content_content_claveTextBox").SendKeys password
        .FindElementById("content_content_iniciarSesionButton").Click
         
        
        Set listDrow = g.FindElementById("content_content_vistaDropDownList").AsSelect
        listDrow.SelectByValue view

        Application.wait Now() + TimeValue("00:00:02")
        
        .FindElementById("content_content_fechaTextBox").Clear
        Application.wait Now() + TimeValue("00:00:01")
        .FindElementById("content_content_fechaTextBox").SendKeys date_view
        .FindElementByXPath("/html/body/form/div[5]/div/div/div[3]/div[2]/div[1]/div/div[4]/div/span/i").Click
        Application.wait Now() + TimeValue("00:00:01")
        .FindElementById("content_content_generarButton").Click
       
        .FindElementById("content_content_descargarButton").Click
        Application.wait Now() + TimeValue("00:00:02")
        
   
        .Close
    End With

    MsgBox "Descarga de archivo", vbInformation + vbOKOnly
    Exit Sub
Cath:
    MsgBox Err.Description, vbCritical + vbOKOnly, Err.Number
End Sub

Sub view_client()
    
    Rem HOLGURAS
    Rem Este proceso mediante web scraping usando SeleniumBasic, realizamos dar una extensión de plazo al los clientes
    Rem cada sierto posiblemente necesite actualizar el chromedriver.exe,dado que se actualiza el navegador también
    Rem la ruta posible donde se encuentra el chromedriver C:\Users\MI_USUARIO\AppData\Local\SeleniumBasic
    
    Rem este procedimiento esta asociado al botón de la ribbon "Prorroga para clientes"
    
    Dim g As New WebDriver
    Dim url As String
    
    '''''
    Dim cedula As String
    Dim orden As Long
    ''''
    Dim temaSelect As SelectElement
    Dim catSelect As SelectElement
    Dim asunSelect As SelectElement
    Dim areaResp As SelectElement
    Dim elemProcessComplete As Boolean
    Dim elem2 As WebElement
       
    On Error GoTo Cath

    url = "http://ombu/SlxClient/Contact.aspx"
    cedula = "1052406961"
    orden = 17428831
    
    With g
        .Start "chrome"
        .Get url
        
        .FindElementById("ctl00_ContentPlaceHolderArea_slxLogin_UserName").SendKeys user
        .FindElementById("ctl00_ContentPlaceHolderArea_slxLogin_Password").SendKeys password
        .FindElementById("ctl00_ContentPlaceHolderArea_slxLogin_btnLogin").Click
        
        Rem buscamos por cedula
        .FindElementById("ext-gen62").Click
        .FindElementById("value_0").SendKeys cedula
        .FindElementByXPath("/html/body/div[13]/div[2]/div[1]/div/div/div/div/div/div/div/div/div/input[2]").Click
        
        .wait 1000
'        Set elem2 = .FindElementByXPath("//*[@id='MainList_list_row_1']/table/tbody/tr/td[4]/div")
        Set elem2 = .FindElementByXPath("/html/body/form/div[4]/div/div/div[3]/div/div/div[1]/div/div[2]/div[1]/div/div/div[1]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[4]/div/a[2]")
        elem2.FindElementByXPath("//*[@id='MainList_list_row_1']/table/tbody/tr/td[4]/div/a[2]").Click
        
        .FindElementByXPath("//*[@id='ctl00_TabControl_element_AccountTickets_element_view_AccountTickets_AccountTickets_btnAddTicket']").Click
        .FindElementByXPath("//*[@id='ctl00_MainContent_InsertTicket_dplArea_Area_ShowBtn']").Click
        
    Stop
        Rem selección de asuntos
        .wait 1000
       Set temaSelect = .FindElementByXPath("/html/body/form/div[4]/div/div/div[3]/div/div/div/div/div/div/div[2]/div/table/tbody/tr[3]/td[1]/div/div/div[4]/div[1]/div/table/tbody/tr[2]/td[1]/select").AsSelect
       temaSelect.SelectByText "Suscripción"
       
       .wait 1000
       Set catSelect = .FindElementByXPath("/html/body/form/div[4]/div/div/div[3]/div/div/div/div/div/div/div[2]/div/table/tbody/tr[3]/td[1]/div/div/div[4]/div[1]/div/table/tbody/tr[2]/td[2]/select").AsSelect
       catSelect.SelectByText "Pedido Interno"

        .wait 1000
       Set asunSelect = .FindElementByXPath("/html/body/form/div[4]/div/div/div[3]/div/div/div/div/div/div/div[2]/div/table/tbody/tr[3]/td[1]/div/div/div[4]/div[1]/div/table/tbody/tr[2]/td[3]/select").AsSelect
       asunSelect.SelectByValue "Q6UJ9BWDOYZG"
       .FindElementById("ctl00_MainContent_InsertTicket_dplArea_btnOk").Click

        Rem despues de guardar
        Rem seleccionamos POR HOLGURA
        .FindElementById("ctl00_MainContent_InsertTicket_cmdSave").Click
        
        Rem DIAGNOSTICO
        .FindElementById("ctl00_MainContent_TicketDetails_Diagnostico_LookupBtn").Click
        .FindElementsByCss(".x-grid3-cell-inner.x-grid3-col-0")(1).ClickDouble

        Rem seleccinamos por ADMINISTRATIVO
        adm = "Administrativo"
        .FindElementById("ctl00_MainContent_TicketDetails_AreaResponsable_Text").SendKeys adm
        
        Rem seleccionamos (ajuste inferior a 30 dias)
        .FindElementById("ctl00_MainContent_TicketDetails_Solucion_LookupBtn").Click
        .FindElementsByCss(".x-grid3-cell-inner.x-grid3-col-0")(3).ClickDouble
        Stop
        
        
        Rem ingresamos orden------------------------------------------------------------------------------------
        .FindElementById("ctl00_MainContent_TicketClubDetails_ETOrden_LookupBtn").Click
        .FindElementById("ctl00_MainContent_TicketClubDetails_ETOrden_lookup_values_input").SendKeys orden
        .FindElementById("ext-gen734").Click
        
        Dim ordenes As WebElements
        
        Set ordenes = .FindElementsByCss(".x-grid3-cell-inner.x-grid3-col-0")
        
        elemProcessComplete = executeWebelement(ordenes, orden)
        
        If Not elemProcessComplete Then Exit Sub 'se cierra el proceso si no completo la acción de seleccionar dicho elemento
        Rem fin del proceso de orden------------------------------------------------------------------------------
           
Stop
    End With
    Exit Sub
Cath:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
End Sub

Public Function executeWebelement(ByRef elements As WebElements, ByVal valueCompare As String) As Boolean
    
    On Error GoTo Cath
    
    For Each elem In elements
        If elem.Text = valueCompare Then
            elem.ClickDouble
            executeWebelement = True
            Exit For
        End If
    Next elem
    
Exit Function
Cath:
    Debug.Print Err.Description
    Debug.Print Err.Number
    Debug.Print "No se completo la acción para este valor "; valueCompare
End Function

Sub join_all_sheets(control As IRibbonControl)
    
    Rem Este rutina se encarga de unir los libros de una determinada carpeta, consolidando en una hoja nueva
    Rem con el nombre "consolidado", este proceso copia los encabezados de una hoja oculta.
    
    Rem este procedimiento esta asociado al botón de la ribbon "Unir hojas"
    
    Dim count As Long
    Dim rs As ADODB.Recordset
    Dim table As String
    Dim source As String
    Dim directory As String
    Dim file As String
    Dim c As Integer
    Dim t As Integer
    Dim date_download_file As Date
    
    Dim fso As New Scripting.FileSystemObject
    Dim f As file
    Dim fr As folder
    
    On Error GoTo Cath
    
    Call deleteSheet_copyHeader
    
    Application.ScreenUpdating = False
   
    directory = InputBox("Ingrese una ruta" & vbCrLf & vbCrLf & "Ejm:" & vbCrLf & "C:\Users\gato\Desktop", "Unir libros")
    Set fr = fso.GetFolder(directory)
    
    For Each f In fr.Files
        count = ThisWorkbook.Sheets(NAME_HJ).Range("a1").CurrentRegion.Rows.count + 1
    
        source = directory & "\" & f.Name
        date_download_file = getDate(f.Name)
        
        Workbooks.Open source
        table = ActiveWorkbook.ActiveSheet.Name
    
        Set rs = getRecordset(table, source)
        
        If Not rs Is Nothing Then
            ThisWorkbook.Sheets(NAME_HJ).Range("B" & count).CopyFromRecordset rs
            ActiveWorkbook.Close
            Call setDateDownload(NAME_HJ, date_download_file)
            c = c + 1
        End If
        
        ThisWorkbook.Save
        t = t + 1
    Next f
    Application.ScreenUpdating = True
    
    MsgBox "Número de libros copiados " & c & " de " & t, vbInformation, "Copia de libros"
    Exit Sub
Cath:
    Debug.Print Err.Number & vbCrLf; Err.Description
    MsgBox Err.Description, vbCritical, "ERROR -- " & Err.Number
    
End Sub
