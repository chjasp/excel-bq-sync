'====================
' UI SECTION
'====================

' Subroutine to add Download and Upload buttons to the "Controls" sheet
Public Sub AddStyledButtons()
    Dim ws As Worksheet
    Dim downloadBtn As Shape
    Dim uploadBtn As Shape
    
    Set ws = ThisWorkbook.Worksheets("Controls")
    
    ' Check if buttons already exist and delete them
    On Error Resume Next
    ws.Shapes("DownloadButton").Delete
    ws.Shapes("UploadButton").Delete
    On Error GoTo 0
    
    ' Add the Download button as a shape
    Set downloadBtn = ws.Shapes.AddShape(msoShapeRectangle, 100, 170, 120, 40)
    With downloadBtn
        .Name = "DownloadButton"
        .OnAction = "DownloadButtonClick"
        .TextFrame.Characters.Text = "Download"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .Fill.ForeColor.RGB = RGB(42, 93, 203) ' Set background color
        .Line.Visible = msoFalse ' Remove border if desired
    End With
    
    ' Add the Upload button as a shape
    Set uploadBtn = ws.Shapes.AddShape(msoShapeRectangle, 230, 170, 120, 40)
    With uploadBtn
        .Name = "UploadButton"
        .OnAction = "UploadButtonClick"
        .TextFrame.Characters.Text = "Upload"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .Fill.ForeColor.RGB = RGB(42, 93, 203) ' Set background color
        .Line.Visible = msoFalse ' Remove border if desired
    End With
    
    MsgBox "Styled Download and Upload buttons added to the Controls sheet.", vbInformation
End Sub


'====================
' DOWNLOAD SECTION
'====================

' Subroutine triggered by the Download button
Public Sub DownloadButtonClick()
    ExecuteBigQueryAndPopulateSheet
End Sub

' Subroutine to query BigQuery and add results to the "Data" sheet
Public Sub ExecuteBigQueryAndPopulateSheet()
    Dim jwt As String
    Dim queryResult As String
    Dim sqlQuery As String
    Dim ws As Worksheet
    
    Debug.Print "Starting ExecuteBigQueryAndPopulateSheet"
    
    ' Get values from the Controls sheet
    Set ws = ThisWorkbook.Worksheets("Controls")
    sqlQuery = ws.Range("C6").Value
    Debug.Print "SQL Query: " & sqlQuery
    
    Debug.Print "Getting JWT..."
    jwt = GetSignedJWT()
    Debug.Print "JWT received: " & Left(jwt, 20) & "..." ' Print first 20 characters for security
    
    Debug.Print "Querying BigQuery..."
    queryResult = QueryBigQuery(jwt, sqlQuery)
    Debug.Print "Query Result received. Adding to 'Data' sheet..."
    
    ' Add the query result to the "Data" sheet
    AddQueryResultToDataSheet queryResult
    
    Debug.Print "ExecuteBigQueryAndPopulateSheet completed. Data added to 'Data' sheet."
    MsgBox "Data download completed.", vbInformation
End Sub

' Function to query BigQuery using the provided access token and SQL query
Public Function QueryBigQuery(accessToken As String, query As String) As String
    Dim bigQueryURL As String
    Dim requestBody As String
    Dim response As String
    
    Debug.Print "QueryBigQuery: Starting"
    
    bigQueryURL = "https://bigquery.googleapis.com/bigquery/v2/projects/main-dev-431619/queries"
    requestBody = "{""query"": """ & query & """, ""useLegacySql"": false}"
    
    Debug.Print "BigQuery URL: " & bigQueryURL
    Debug.Print "Request Body: " & requestBody
    
    ' Send the query to BigQuery
    response = MacHttpRequestWithAuth(bigQueryURL, requestBody, accessToken)
    
    Debug.Print "Response:" & response
    
    QueryBigQuery = response ' You may want to parse this response further
    Debug.Print "QueryBigQuery: Completed"
End Function

' Subroutine to add query results to the "Data" sheet
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

'====================
' UPLOAD SECTION
'====================

' Subroutine triggered by the Upload button
Public Sub UploadButtonClick()
    UploadDataToBigQuery
End Sub

' Subroutine to upload data to BigQuery
Public Sub UploadDataToBigQuery()
    Dim jwt As String
    Dim projectId As String
    Dim datasetName As String
    Dim tableName As String
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim uploadResult As String
    Dim parsedDate As Date
    
    Debug.Print "Starting UploadDataToBigQuery"
    
    ' Get values from the Controls sheet
    Set ws = ThisWorkbook.Worksheets("Controls")
    projectId = ws.Range("C3").Value
    datasetName = ws.Range("C4").Value
    tableName = ws.Range("C5").Value
    
    ' Get the Data sheet
    Set dataWs = ThisWorkbook.Worksheets("Data")
    
    ' Get JWT
    Debug.Print "Getting JWT..."
    jwt = GetSignedJWT()
    Debug.Print "JWT received: " & Left(jwt, 20) & "..." ' Print first 20 characters for security
    
    ' Prepare and upload data using DML statements
    Debug.Print "Preparing and uploading data..."
    uploadResult = PrepareAndUploadData(jwt, projectId, datasetName, tableName, dataWs)

    
    Debug.Print "Upload Result: " & uploadResult
    MsgBox "Data upload completed.", vbInformation
End Sub

' Function to upload data to BigQuery using the insertAll endpoint
Private Function PrepareAndUploadData(jwt As String, projectId As String, datasetName As String, tableName As String, dataWs As Worksheet) As String
    Dim dataRange As Range
    Dim headers As Variant
    Dim data As Variant
    Dim i As Long, j As Long
    Dim jsonRows As String
    Dim requestBody As String
    Dim bigQueryURL As String
    Dim response As String
    Dim batchSize As Long
    Dim currentBatchSize As Long

    ' Determine the data range
    Set dataRange = dataWs.UsedRange

    ' Ensure there's at least one row for headers
    If dataRange.Rows.Count < 2 Then
        MsgBox "No data found in the 'Data' sheet.", vbExclamation
        PrepareAndUploadDataInsertAll = ""
        Exit Function
    End If

    ' Get headers and data
    headers = dataRange.Rows(1).Value ' Headers are in the first row
    data = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count).Value

    ' Set the BigQuery insertAll URL
    bigQueryURL = "https://bigquery.googleapis.com/bigquery/v2/projects/" & projectId & "/datasets/" & datasetName & "/tables/" & tableName & "/insertAll"

    ' Initialize variables
    batchSize = 500  ' Adjust based on your requirements and limits
    currentBatchSize = 0
    jsonRows = ""

    ' Loop through the data rows and construct JSON rows
    For i = 1 To UBound(data, 1)
        Dim jsonRow As String
        jsonRow = "{""json"":{"

        For j = 1 To UBound(data, 2)
            Dim columnName As String
            Dim cellValue As String

            columnName = headers(1, j)
            cellValue = data(i, j)

            ' Handle null values
            If IsError(cellValue) Or IsEmpty(cellValue) Or cellValue = "" Then
                cellValue = ""
            Else
                cellValue = CStr(cellValue)
                ' Escape double quotes
                cellValue = Replace(cellValue, """", "\""")
            End If

            ' Append to jsonRow
            If j > 1 Then jsonRow = jsonRow & ","
            jsonRow = jsonRow & """" & columnName & """:""" & cellValue & """"
        Next j
      
        ' Add ingestion date
        jsonRow = jsonRow & ",""ingestion_dt"":""" & Format(Now, "yyyy-MM-dd") & """"

        ' Add ingestion timestamp
        jsonRow = jsonRow & ",""ingestion_ts"":""" & Format(Now, "yyyy-MM-dd HH:mm:ss") & """"
        
        jsonRow = jsonRow & "}}"

        ' Append to jsonRows
        If jsonRows <> "" Then jsonRows = jsonRows & ","
        jsonRows = jsonRows & jsonRow

        currentBatchSize = currentBatchSize + 1

        ' If batch size is reached or last row, send the request
        If currentBatchSize >= batchSize Or i = UBound(data, 1) Then
            requestBody = "{" & """rows"":[" & jsonRows & "]}"
            Debug.Print "Request Body: " & requestBody

            ' Send the request
            response = MacHttpRequestWithAuth(bigQueryURL, requestBody, jwt)

            ' Handle the response as needed
            Debug.Print "Response: " & response

            ' Reset for the next batch
            jsonRows = ""
            currentBatchSize = 0
        End If
    Next i

    PrepareAndUploadData = "Data upload completed."
End Function

'====================
' UTILITY FUNCTIONS
'====================

' Function to read service account info from the "Controls" sheet
Private Function ReadServiceAccountInfo() As String
    Dim ws As Worksheet
    Dim keyContent As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Controls")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Controls' not found!", vbExclamation
        Exit Function
    End If
    
    ' Assuming the key content is in cell B1 of the "Controls" sheet
    keyContent = ws.Range("C2").Value
    
    ' Remove any leading/trailing whitespace
    keyContent = Trim(keyContent)
    
    ' Replace actual newline characters with \n
    keyContent = Replace(keyContent, "\n", "\\n")
    
    ' Ensure the content is properly formatted as JSON
    If Left(keyContent, 1) <> "{" Or Right(keyContent, 1) <> "}" Then
        keyContent = "{" & keyContent & "}"
    End If
    
    ReadServiceAccountInfo = keyContent
End Function

' Function to get a signed JWT from a cloud function
Public Function GetSignedJWT() As String
    Dim cloudFunctionURL As String
    Dim response As String
    
    Debug.Print "GetSignedJWT: Starting"
    
    cloudFunctionURL = "https://europe-west3-main-dev-431619.cloudfunctions.net/jwt-creator"
    Debug.Print "Cloud Function URL: " & cloudFunctionURL
    
    ' Call the cloud function to get the JWT
    response = MacHttpRequestWithAuth(cloudFunctionURL, "{""service_account_info"": " & ReadServiceAccountInfo() & "}")
    
    Debug.Print "RESPONSE:"
    Debug.Print response
   
    ' Extract the JWT from the response
    GetSignedJWT = ParseJWTFromResponse(response)
    Debug.Print "GetSignedJWT: Completed"
End Function

' Function to make HTTP requests with optional authentication
Private Function MacHttpRequestWithAuth(url As String, requestBody As String, Optional accessToken As String = "", Optional contentType As String = "application/json") As String
    Dim tempFilePath As String
    Dim curlCmd As String
    Dim appleScriptCmd As String
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
    appleScriptCmd = "do shell script """ & Replace(curlCmd, """", "\""") & """"
    Debug.Print "SCRIPT COMMAND:"
    Debug.Print appleScriptCmd
    
    ' Execute the curl command using MacScript
    On Error GoTo MacScriptError
    MacScript appleScriptCmd
    
    ' Read the response from the temporary file
    On Error GoTo ReadError
    Open tempFilePath For Input As #1
    MacHttpRequestWithAuth = Input$(LOF(1), 1)
    Close #1
    
    ' Delete the temporary file
    On Error Resume Next
    Kill tempFilePath
    On Error GoTo 0
    
    Exit Function

' Error Handlers
MacScriptError:
    MsgBox "Error executing shell script: " & Err.Description, vbCritical
    MacHttpRequestWithAuth = ""
    Exit Function

FileError:
    MsgBox "Error creating temporary file: " & Err.Description, vbCritical
    MacHttpRequestWithAuth = ""
    Exit Function

ReadError:
    MsgBox "Error reading response from temporary file: " & Err.Description, vbCritical
    MacHttpRequestWithAuth = ""
    Exit Function
End Function

' Function to parse the JWT from the cloud function response
Private Function ParseJWTFromResponse(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(response, """jwt"":") + 7
    endPos = InStr(startPos, response, """") - 1
    
    ParseJWTFromResponse = Mid(response, startPos, endPos - startPos + 1)
End Function