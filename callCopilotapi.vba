Sub CallCopilotAPI()
    Dim wsData As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim userPrompt As String
    Dim jsonPayload As String
    Dim http As Object
    Dim apiURL As String
    Dim response As String
    
    ' 1. Get the prompt from your Native Text Box Form Control
    ' Note: Adjust "Text Box 1" to match the name Excel gave your control
    On Error Resume Next
    userPrompt = ActiveSheet.TextBoxes("myprompt").Text
    On Error GoTo 0
    
    If Trim(userPrompt) = "" Then
        MsgBox "Please enter a question first.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ' 2. Target the specific compliance-approved data tab
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("databricks usage cost")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "Tab 'databricks usage cost' not found!", vbCritical, "Error"
        Exit Sub
    End If
    
    ' 3. Dynamically capture the used data grid
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    ' 4. Format a simple JSON Payload (Safe for corporate proxies)
    ' This formats data cell-by-cell to ensure no external JSON parsing libraries are needed
    jsonPayload = "{""prompt"":""" & Replace(userPrompt, """", "\""") & """,""data"":["
    
    Dim r As Long, c As Long
    For r = 1 To lastRow
        jsonPayload = jsonPayload & "["
        For c = 1 To lastCol
            Dim cellValue As String
            cellValue = Replace(wsData.Cells(r, c).Text, """", "\""")
            jsonPayload = jsonPayload & """" & cellValue & """"
            If c < lastCol Then jsonPayload = jsonPayload & ","
        Next c
        jsonPayload = jsonPayload & "]"
        If r < lastRow Then jsonPayload = jsonPayload & ","
    Next r
    jsonPayload = jsonPayload & "]}"
    
    ' 5. Send secure API Call using native Windows HTTP Services
    apiURL = "https://your-bank-internal-copilot-proxy/api/v1" ' Replace with your bank's approved internal gateway
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", apiURL, False
    http.setRequestHeader "Content-Type", "application/json"
    ' http.setRequestHeader "Authorization", "Bearer YOUR_BANK_TOKEN" ' Uncomment if required by infrastructure
    
    On Error GoTo HttpError
    http.send jsonPayload
    response = http.responseText
    
    ' 6. Output the raw text response back safely to the user
    MsgBox "Copilot Response:" & vbCrLf & vbCrLf & response, vbInformation, "Success"
    Exit Sub

HttpError:
    MsgBox "API connection blocked or failed. Please check network/proxy rules.", vbCritical, "Compliance Connection Error"
End Sub


