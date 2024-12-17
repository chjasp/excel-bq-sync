Public Sub ExecuteBigQueryAndPopulateSheet()
    Dim rawAccountInfo As String
    Dim parsedAccountInfo As String
    Dim accessToken As String
    Dim queryResult As String
    Dim sqlQuery As String
    Dim ws As Worksheet
    
    Debug.Print "Starting ExecuteBigQueryAndPopulateSheet"

    Set ws = ThisWorkbook.Worksheets("Controls")
    rawAccountInfo = ws.Range("B1").Value
    parsedAccountInfo = ParseServiceAccountInfo(rawAccountInfo)
    sqlQuery = ws.Range("B5").Value
    
    Debug.Print "Getting AccessToken..."
    accessToken = GetAccessToken(parsedAccountInfo)
    
    Debug.Print "Querying BigQuery..."
    queryResult = QueryBigQuery(accessToken, sqlQuery)

    Debug.Print "Query Result received. Adding to 'Data' sheet..."
    AddQueryResultToDataSheet queryResult
    
    Debug.Print "ExecuteBigQueryAndPopulateSheet completed. Data added to 'Data' sheet."
    MsgBox "Data download completed.", vbInformation
End Sub

Public Function QueryBigQuery(accessToken As String, query As String) As String
    Dim bigQueryURL As String
    Dim requestBody As String
    Dim response As String
    
    Debug.Print "QueryBigQuery: Starting"
    
    bigQueryURL = "https://bigquery.googleapis.com/bigquery/v2/projects/PROJECT_ID/queries"
    requestBody = "{""query"": """ & query & """, ""useLegacySql"": false}"
    
    Debug.Print "BigQuery URL: " & bigQueryURL
    Debug.Print "Request Body: " & requestBody
    
    ' Send the query to BigQuery
    response = HttpRequest(bigQueryURL, requestBody, accessToken)
    
    Debug.Print "Response:" & response
    
    QueryBigQuery = response ' You may want to parse this response further
    Debug.Print "QueryBigQuery: Completed"
End Function

Private Sub AddQueryResultToDataSheet(queryResult As String)
    Dim ws As Worksheet
    Dim json As Object
    Dim schema As Object
    Dim fields As Variant
    Dim rows As Variant
    Dim i As Long, j As Long
    Dim field As Variant
    Dim row As Variant
    
    ' Parse the JSON response
    On Error GoTo JsonParseError
    Set json = JsonConverter.ParseJson(queryResult)
    
    ' Create or get the "Data" sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Data"
    End If
    
    ' Clear existing content
    ws.Cells.Clear
    
    ' Extract and write headers
    Set schema = json("schema")
    Set fields = schema("fields")
    j = 1
    For Each field In fields
        ws.Cells(1, j).Value = field("name")
        j = j + 1
    Next field
    
    ' Check if there are any rows
    If json.Exists("rows") Then
        ' Extract and write rows
        Set rows = json("rows")
        For i = 1 To rows.Count
            Set row = rows(i)("f")
            For j = 1 To row.Count
                ws.Cells(i + 1, j).Value = row(j)("v")
            Next j
        Next i
        
        Debug.Print "Data added to 'Data' sheet successfully."
        Debug.Print "Rows added: " & rows.Count
    Else
        ws.Cells(2, 1).Value = "No data found"
        Debug.Print "No data found in the query result."
    End If
    
    Debug.Print "Columns added: " & fields.Count
    
    ' Autofit columns
    ws.Columns.AutoFit
    
    Exit Sub

JsonParseError:
    MsgBox "Error parsing query result: " & Err.Description, vbCritical
End Sub

Private Function ParseServiceAccountInfo(rawAccountInfo As String) As String
    Dim parsedAccountInfo As String
    
    parsedAccountInfo = Trim(rawAccountInfo)
    parsedAccountInfo = Replace(parsedAccountInfo, "\n", "\\n")

    If Left(parsedAccountInfo, 1) <> "{" Or Right(parsedAccountInfo, 1) <> "}" Then
        parsedAccountInfo = "{" & parsedAccountInfo & "}"
    End If
    
    ParseServiceAccountInfo = parsedAccountInfo
End Function


Public Function GetAccessToken(accountInfo As String) As String
    Dim cloudFunctionURL As String
    Dim response As String
    
    Debug.Print "GetAccessToken: Starting process"
    
    cloudFunctionURL = ""
    
    Debug.Print "  Calling cloud function at: " & cloudFunctionURL
    
    ' Call the cloud function to get the Access Token
    response = HttpRequest(cloudFunctionURL, "{""service_account_info"": " & accountInfo & "}")
   
    ' Extract the Access Token from the response
    GetAccessToken = ParseAccessTokenFromResponse(response)
    
    Debug.Print "GetAccessToken: Process completed successfully"
End Function

Private Function ParseAccessTokenFromResponse(response As String) As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    ' Parse the JSON response
    Set json = JsonConverter.ParseJson(response)
    
    ' Extract the access token
    If json.Exists("access_token") Then
        ParseAccessTokenFromResponse = json("access_token")
    Else
        Debug.Print "Access token not found in response"
        ParseAccessTokenFromResponse = ""
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "Error in ParseAccessTokenFromResponse: " & Err.Description
    ParseAccessTokenFromResponse = ""
End Function

Private Function HttpRequest(url As String, requestBody As String, Optional accessToken As String = "", Optional contentType As String = "application/json") As String
    Dim tempFilePath As String
    Dim curlCmd As String
    Dim scriptCmd As String
    Dim headers As String
    
    tempFilePath = Environ("TMPDIR") & "temp_response.txt"
    
    ' Create a temporary file to store the response
    On Error GoTo FileError
    Open tempFilePath For Output As #1
    Close #1
    
    ' Set the Content-Type header dynamically
    headers = "-H ""Content-Type: " & contentType & """"
    If accessToken <> "" Then
        headers = headers & " -H ""Authorization: Bearer " & accessToken & """"
    End If
    
    ' Construct the curl command
    curlCmd = "curl -X POST " & headers & " -d '" & requestBody & "' " & url & " > " & tempFilePath
    scriptCmd = "do shell script """ & Replace(curlCmd, """", "\""") & """"
    
    ' Execute the curl command using MacScript
    On Error GoTo MacScriptError
    MacScript scriptCmd
    
    ' Read the response from the temporary file
    On Error GoTo ReadError
    Open tempFilePath For Input As #1
    HttpRequest = Input$(LOF(1), 1)
    Close #1
    
    ' Delete the temporary file
    On Error Resume Next
    Kill tempFilePath
    On Error GoTo 0
    
    Exit Function

' Error Handlers
MacScriptError:
    MsgBox "Error executing shell script: " & Err.Description, vbCritical
    HttpRequest = ""
    Exit Function

FileError:
    MsgBox "Error creating temporary file: " & Err.Description, vbCritical
    HttpRequest = ""
    Exit Function

ReadError:
    MsgBox "Error reading response from temporary file: " & Err.Description, vbCritical
    HttpRequest = ""
    Exit Function
End Function

' Subroutine to upload data to BigQuery
Public Sub UploadDataToBigQuery()
    Dim accessToken As String
    Dim projectId As String
    Dim datasetName As String
    Dim tableName As String
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim uploadResult As String
    Dim rawAccountInfo As String
    Dim parsedAccountInfo As String
    
    Debug.Print "Starting UploadDataToBigQuery"
    
    ' Get values from the Controls sheet
    Set ws = ThisWorkbook.Worksheets("Controls")
    rawAccountInfo = ws.Range("B1").Value
    parsedAccountInfo = ParseServiceAccountInfo(rawAccountInfo)
    projectId = ws.Range("B2").Value
    datasetName = ws.Range("B3").Value
    tableName = ws.Range("B4").Value
    
    ' Get the Data sheet
    Set dataWs = ThisWorkbook.Worksheets("Data")
    
    ' Get Access Token
    Debug.Print "Getting Access Token..."
    accessToken = GetAccessToken(parsedAccountInfo)
    Debug.Print "Access Token received: " & Left(accessToken, 20) & "..."
    
    ' Prepare and upload data using DML statements
    Debug.Print "Preparing and uploading data..."
    uploadResult = PrepareAndUploadData(accessToken, projectId, datasetName, tableName, dataWs)
    
    Debug.Print "Upload Result: " & uploadResult
    MsgBox "Data upload completed.", vbInformation
End Sub

' Function to upload data to BigQuery using the insertAll endpoint
Private Function PrepareAndUploadData(accessToken As String, projectId As String, datasetName As String, tableName As String, dataWs As Worksheet) As String
    Dim dataRange As Range
    Dim headers As Variant
    Dim data As Variant
    Dim i As Long, j As Long
    Dim jsonRows As String
    Dim requestBody As String
    Dim bigQueryURL As String
    Dim response As String

    ' Determine the data range
    Set dataRange = dataWs.UsedRange

    ' Ensure there's at least one row for headers
    If dataRange.Rows.Count < 2 Then
        MsgBox "No data found in the 'Data' sheet.", vbExclamation
        PrepareAndUploadData = ""
        Exit Function
    End If

    ' Get headers and data
    headers = dataRange.Rows(1).Value ' Headers are in the first row
    data = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count).Value

    ' Set the BigQuery insertAll URL
    bigQueryURL = "https://bigquery.googleapis.com/bigquery/v2/projects/" & projectId & "/datasets/" & datasetName & "/tables/" & tableName & "/insertAll"

    ' Initialize jsonRows
    jsonRows = ""

    ' Loop through the data rows and construct JSON rows
    For i = 1 To UBound(data, 1)
        Dim jsonRow As String
        jsonRow = "{""json"":{"

        For j = 1 To UBound(data, 2)
            Dim columnName As String
            Dim cellValue As Variant

            columnName = headers(1, j)
            cellValue = data(i, j)

            ' Handle different data types
            If j > 1 Then jsonRow = jsonRow & ","
            
            jsonRow = jsonRow & """" & columnName & """:"
            
            ' Handle null values
            If IsError(cellValue) Or IsEmpty(cellValue) Or cellValue = "" Then
                jsonRow = jsonRow & "null"
            ' Handle numbers
            ElseIf IsNumeric(cellValue) Then
                jsonRow = jsonRow & CStr(cellValue)
            ' Handle booleans
            ElseIf VarType(cellValue) = vbBoolean Then
                jsonRow = jsonRow & LCase(CStr(cellValue))
            ' Handle strings
            Else
                cellValue = Replace(CStr(cellValue), """", "\""")
                cellValue = Replace(cellValue, vbNewLine, "\n")
                jsonRow = jsonRow & """" & cellValue & """"
            End If
        Next j
        
        jsonRow = jsonRow & "}}"

        ' Append to jsonRows
        If jsonRows <> "" Then jsonRows = jsonRows & ","
        jsonRows = jsonRows & jsonRow
    Next i

    ' Send all rows in a single request
    requestBody = "{" & """rows"":[" & jsonRows & "]}"
    Debug.Print "Request Body: " & requestBody

    ' Send the request
    response = HttpRequest(bigQueryURL, requestBody, accessToken)
    Debug.Print "Response: " & response

    PrepareAndUploadData = "Data upload completed."
End Function