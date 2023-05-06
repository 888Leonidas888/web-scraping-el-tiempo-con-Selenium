Attribute VB_Name = "packageUDFCedula"
Const CAP_COLOMBIA = "¿ Cual es la Capital de Colombia (sin tilde)?"
Const CAP_ATLANTICO = "¿ Cual es la Capital del Atlantico?"
Const CAP_CAUCA = "¿ Cual es la Capital del Vallle del Cauca?"
Const CAP_ANTIOQUIA = "¿ Cual es la Capital de Antioquia (sin tilde)?"
Const CUANTO_ES = "¿ Cuanto es "
Const PRIM_NOMBRE = "¿Cual es el primer nombre de la persona a la cual esta expidiendo el certificado?"
Const PRIM_LETRAS_NOMBRE = "¿Escriba las dos primeras letras del primer nombre de la persona a la cual esta expidiendo el certificado?"
Const DOS_ULTIM_DIG_DOCUMENTO = "¿Escriba los dos ultimos digitos del documento a consultar?"
Const TRES_PRIM_DIG_DOCUMENTO = "¿Escriba los tres primeros digitos del documento a consultar?"
Const CANT_LETRAS_PRIMER_NOMBRE = "¿Escriba la cantidad de letras del primer nombre de la persona a la cual esta expidiendo el certificado?"

Type fullName
    name As String
    lastName As String
End Type

Function searchPersonForCedula(ByVal cedula As String) As String
    
    Dim c As New ChromeDriver
    Dim url As String
    Dim drownList As SelectElement
    Dim spans As WebElements
    Dim div As WebElement
    Dim getQuestion As String
    Dim getAnswer As Variant
    Dim nextCell As Integer
    Dim recaptcha As Boolean
    Dim noFound As String
    Dim fullName As String
    
    On Error GoTo Cath
    
    url = "https://apps.procuraduria.gov.co/webcert/inicio.aspx?tpo=1\"
    c.Start "chrome"
    c.Get url
    
    Set drownList = c.FindElementById("ddlTipoID").AsSelect
    drownList.SelectByValue 1
    
    c.FindElementById("txtNumID").SendKeys cedula
    
    Do
        recaptcha = False
        Rem recupera pregunta
        getQuestion = c.FindElementById("lblPregunta").text
        
        Rem procesa la pregunta
        getAnswer = answerCaptcha(getQuestion, cedula, recaptcha)
        
        Rem devuelva la respuesta en el html
        c.FindElementById("txtRespuestaPregunta").SendKeys getAnswer
        
        If recaptcha Then
            c.FindElementById("ImageButton1").Click
            c.Wait 1000
        End If
        
    Loop While recaptcha

    c.FindElementById("btnConsultar").Click

    c.Wait 2000
    
    Set spans = c.FindElementsByCss("#divSec span")
    Set div = c.FindElementByCss("#ValidationSummary1")
    noFound = Trim(div.text)
    
    
    Rem---------------
    Rem separamos nombre de apellidos
    Dim surnames As String
    Dim names As String

    surnames = spans.Item(spans.Count - 1).text & " " & spans.Item(spans.Count).text
    
    For i = 1 To spans.Count - 2
        names = names & " " & spans.Item(i).text
    Next i
        
    names = Trim(names)
    
    fullName = names & "," & surnames
    Rem---------------
    
'    For Each span In spans
''        Debug.Print span.Text
'        fullName = fullName & " " & span.text
'    Next span
    
'    MsgBox c.FindElementByClass("datosConsultado").Text
'    MsgBox fullName
    c.Close
    
    If fullName <> Empty Then
        searchPersonForCedula = Trim(fullName)
    ElseIf noFound <> Empty Then
        searchPersonForCedula = "No registrado"
    Else
        searchPersonForCedula = "Error"
    End If
    
    Exit Function
    
Cath:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
    searchPersonForCedula = "Error"
    
End Function
Function answer(ByVal question) As String
    Dim r As New RegExp
    Dim rMatch As MatchCollection
    
    With r
        .Global = True
        .Pattern = "\d+|[\+\-\X]"
        
        Set rMatch = .Execute(question)
        
        If rMatch.Count > 0 Then
'            MsgBox rMatch.Item(0)
            For i = 0 To rMatch.Count - 1
                Debug.Print rMatch.Item(i)
            Next i
        End If
    End With
End Function

Function answerCaptcha(ByVal question As String, ByVal cedula As String, ByRef recaptcha As Boolean) As Variant

    Select Case True
        Case InStr(question, CAP_COLOMBIA) > 0
            answerCaptcha = "bogota"
        Case InStr(question, CAP_ATLANTICO) > 0
            answerCaptcha = "barranquilla"
        Case InStr(question, CAP_CAUCA) > 0
            answerCaptcha = "cali"
        Case InStr(question, CAP_ANTIOQUIA) > 0
            answerCaptcha = "medellin"
        Case InStr(question, DOS_ULTIM_DIG_DOCUMENTO) > 0
            answerCaptcha = Right(cedula, 2)
        Case InStr(question, TRES_PRIM_DIG_DOCUMENTO) > 0
            answerCaptcha = Left(cedula, 3)
        Case InStr(question, CUANTO_ES) > 0
            answerCaptcha = answerOperationMath2(question)
        Case InStr(question, CANT_LETRAS_PRIMER_NOMBRE) > 0
            recaptcha = True
        Case InStr(question, PRIM_NOMBRE) > 0
            recaptcha = True
        Case Else
            recaptcha = True
    End Select

End Function

Function answerOperationMath(ByVal question As String) As Variant
    
    Dim r As New RegExp
    Dim rMatch As MatchCollection
    Dim operation As String
    
    With r
        .Global = True
        .Pattern = "\d+ [+-X] \d+"
        
        Set rMatch = .Execute(question)
        
        If rMatch.Count = 1 Then
'            MsgBox rMatch.Item(0)
            answerOperationMath = Replace( _
                                    Replace(rMatch.Item(0), " ", "") _
                                    , "X", "*")
        End If
    End With
    
End Function

Function answerOperationMath2(ByVal question As String) As Variant
    
    Dim r As New RegExp
    Dim rMatch As MatchCollection
    Dim operation As String
    
    Dim num1 As Integer
    Dim num2 As Integer
    Dim operator As String
    
    With r
        .Global = True
        .Pattern = "\d+|[\+\-\X\/]"
        
        Set rMatch = .Execute(question)
        
        If rMatch.Count > 0 Then
            num1 = rMatch.Item(0)
            operator = rMatch.Item(1)
            num2 = rMatch.Item(2)
            
            Select Case operator
                Case Is = "+"
                    answerOperationMath2 = CDbl(num1 + num2)
                Case Is = "-"
                    answerOperationMath2 = CDbl(num1 - num2)
                Case Is = "X"
                    answerOperationMath2 = CDbl(num1 * num2)
                Case Else
                    answerOperationMath2 = CDbl(num1 / num2)
            End Select
        End If
    End With
    
End Function

Sub separate(ByVal text As String)
    
    Dim arrFullName() As String
    Dim fn As fullName
    
    If text <> "Error" Or text <> Empty Then
        arrFullName = Split(text, " ")
        
        fn.lastName = arrFullName(UBound(arrFullName) - 1) & arrFullName(UBound(arrFullName))
    End If
    
End Sub


