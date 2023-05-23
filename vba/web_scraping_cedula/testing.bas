Attribute VB_Name = "Testing"

Sub test1()
    Dim fso As New Scripting.FileSystemObject
    Dim pathDownload As String
    viewText = "Comunicados_Electronicos0 - IE - Tecnologia"
    pathDownload = "C:\Users\dannin\Downloads\" + viewText + "-" + Format(Date, "yyyy-m-dd") + ".xls"
    
        If fso.FileExists(pathDownload) Then
            Debug.Print "existe"
        Else
            Debug.Print "no existe"
        End If
        
'    Do While si
'    Loop
        
End Sub

Sub test2()
    
    
    Dim txt As String
    Dim r As New RegExp
    Dim m As MatchCollection
    
    txt = "Comunicados_Electronicos0 - IE-Tecnologia-2023-3-8.xls"

    With r
        .Global = True
        .Pattern = "\d{1,4}\-\d{1,2}\-\d{1,2}"
        Set m = .Execute(txt)

        Debug.Print CDate(m(0))
        
    End With
End Sub
Sub test3()

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
        .Worksheets("header").Range("A1:XX1").Copy
        With .Worksheets(NAME_HJ)
            .Range("A1").PasteSpecial
            .Range("A1").Select
        End With
    End With
    
    With Application
     .DisplayAlerts = True
     .ScreenUpdating = True
    End With
    
End Sub

Sub test4()
        
    Dim count1 As Long
    Dim count2 As Long
    Dim i As Long
    Dim dDate As Date
    
    dDate = #4/14/2023#
    
    count1 = WorksheetFunction.CountA(Worksheets(NAME_HJ).Range("A:A")) + 1
    count2 = WorksheetFunction.CountA(Worksheets(NAME_HJ).Range("B:B"))
        
    Debug.Print count1
    Debug.Print count2
    
    For i = count1 To count2
        Range("A" & i).Value = dDate
    Next i
    
End Sub

Public Function getRecordsetCount(ByVal table As String, ByVal source As String) As Long
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim strCnn As String
    
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & source & ";Extended Properties=""Excel 8.0;HDR=YES"";"
    sql = "select * from [" & table & "$]"
    
    cnn.ConnectionString = strCnn
    cnn.Open
    
    With rs
        .CursorLocation = adUseClient
        .Open sql, cnn
    End With
    
'    Set rs = cnn.Execute(sql)
'    rs.CursorLocation = adUseClient
    If rs.EOF And rs.BOF Then
        Debug.Print "Sin registros que mostrar para el libro " & source
    Else
'        Set getRecordset = rs
        Debug.Print rs.GetString
        getRecordsetCount = rs.RecordCount
    End If
    
End Function
