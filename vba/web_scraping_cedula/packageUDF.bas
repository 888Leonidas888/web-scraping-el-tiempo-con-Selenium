Attribute VB_Name = "packageUDF"
Sub getRecordset_test()
    'mota
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim strCnn As String
    Dim table As String
    Dim source As String
    
    source = "C:\Users\JHONY\Desktop\juntar hojas\vistas2\Comunicados_Electronicos45 - IE-Tecnologia-2023-2-18.xls"
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & source & ";Extended Properties=""Excel 8.0;HDR=YES"";"
    sql = "select * from [Comunicados_Electronicos45 - IE$]"
    
    cnn.ConnectionString = strCnn
    cnn.Open
    Set rs = cnn.Execute(sql)
    
    If rs.EOF And rs.BOF Then
        Debug.Print "Sin registros que mostrar"
    Else
        Range("B2").CopyFromRecordset rs
        MsgBox "Copia terminada"
        ThisWorkbook.Save
    End If
    
End Sub
Public Function getRecordset(ByVal table As String, ByVal source As String) As ADODB.Recordset

    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim strCnn As String
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & source & ";Extended Properties=""Excel 8.0;HDR=YES"";"
    sql = "select * from [" & table & "$]"
    
    cnn.ConnectionString = strCnn
    cnn.Open
    Set rs = cnn.Execute(sql)
    
    If rs.EOF And rs.BOF Then
        Debug.Print "Sin registros que mostrar para el libro " & source
    Else
        Set getRecordset = rs
    End If
    
End Function
Public Function getDate(ByVal txt As String) As Date
    Dim r As New RegExp
    Dim m As MatchCollection
    Dim d As Date

    With r
        .Global = True
        .Pattern = "\d{1,4}\-\d{1,2}\-\d{1,2}"
        Set m = .Execute(txt)
        
        If m.count = 1 Then
           d = CDate(m(0))
        End If
        
        getDate = d
    End With
    
End Function
Public Sub deleteSheet_copyHeader()

     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    For Each h In Worksheets
        If h.Name = NAME_HJ Then
            Worksheets(NAME_HJ).Delete
            Exit For
        End If
    Next h
    
    Worksheets.Add(before:=Worksheets(Worksheets.count)).Name = NAME_HJ
    
    With ThisWorkbook
        .Worksheets("header").Range("A1:EX1").Copy
        With .Worksheets(NAME_HJ)
            .Range("A1").PasteSpecial
            .Range("A1").Select
        End With
    End With
    
    With Application
     .DisplayAlerts = True
     .ScreenUpdating = True
     .CutCopyMode = False
    End With
    
End Sub
Public Sub setDateDownload(ByVal nameSheet As String, ByVal dDate As Date)
    
    Dim count1 As Long, count2 As Long, i As Long
    
    With WorksheetFunction
        count1 = .CountA(Worksheets(nameSheet).Range("A:A")) + 1
        count2 = .CountA(Worksheets(nameSheet).Range("B:B"))
    End With
    
    For i = count1 To count2
        Range("A" & i).Value = dDate
    Next i
    
End Sub
    




