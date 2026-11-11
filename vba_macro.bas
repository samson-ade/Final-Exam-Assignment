Attribute VB_Name = "LoanPredictionModule"
' =========================================================================================
' MODULE: LoanPredictionModule
' PURPOSE: Integrates Microsoft Excel with FastAPI REST API for Bank Loan Default Prediction
' TARGET ENDPOINT: http://127.0.0.1:8000/predict
' COMPATIBILITY: Microsoft Excel for Windows (MSXML2) & Microsoft Excel for macOS (curl)
' =========================================================================================

Option Explicit

' -----------------------------------------------------------------------------------------
' SUBROUTINE: SubmitLoanPrediction
' PURPOSE: Reads customer attributes from "Single Customer Prediction" sheet,
'          submits JSON payload to FastAPI /predict endpoint, parses the returned
'          probability of default, updates worksheet cells, and presents result alert.
' -----------------------------------------------------------------------------------------
Public Sub SubmitLoanPrediction()
    Dim ws As Worksheet
    Dim apiUrl As String
    Dim jsonPayload As String
    Dim responseText As String
    
    Dim ageVal As Double
    Dim employVal As Double
    Dim addressVal As Double
    Dim debtincVal As Double
    Dim creddebtVal As Double
    Dim othdebtVal As Double
    Dim probDefault As Double
    
    On Error GoTo ErrorHandler
    
    ' Set target worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Single Customer Prediction")
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error GoTo ErrorHandler
    
    ' Validate input cells
    If Not IsNumeric(ws.Range("C6").Value) Or _
       Not IsNumeric(ws.Range("C7").Value) Or _
       Not IsNumeric(ws.Range("C8").Value) Or _
       Not IsNumeric(ws.Range("C9").Value) Or _
       Not IsNumeric(ws.Range("C10").Value) Or _
       Not IsNumeric(ws.Range("C11").Value) Then
        MsgBox "Please ensure all customer financial and demographic fields (C6:C11) contain valid numbers.", _
               vbExclamation, "Input Validation Error"
        Exit Sub
    End If
    
    ' Extract features from customer form
    ageVal = CDbl(ws.Range("C6").Value)
    employVal = CDbl(ws.Range("C7").Value)
    addressVal = CDbl(ws.Range("C8").Value)
    debtincVal = CDbl(ws.Range("C9").Value)
    creddebtVal = CDbl(ws.Range("C10").Value)
    othdebtVal = CDbl(ws.Range("C11").Value)
    
    ' Construct JSON payload
    jsonPayload = "{" & _
        """AGE"":" & Replace(CStr(ageVal), ",", ".") & "," & _
        """EMPLOY"":" & Replace(CStr(employVal), ",", ".") & "," & _
        """ADDRESS"":" & Replace(CStr(addressVal), ",", ".") & "," & _
        """DEBTINC"":" & Replace(CStr(debtincVal), ",", ".") & "," & _
        """CREDDEBT"":" & Replace(CStr(creddebtVal), ",", ".") & "," & _
        """OTHDEBT"":" & Replace(CStr(othdebtVal), ",", ".") & _
        "}"
        
    apiUrl = "http://127.0.0.1:8000/predict"
    
    ' Visual feedback
    ws.Range("C17").Value = "Connecting to API..."
    
    ' Cross-Platform HTTP POST Execution
    #If Mac Then
        ' Execution path for macOS using curl
        Dim macCurlCmd As String
        macCurlCmd = "do shell script ""curl -s -X POST -H \""Content-Type: application/json\"" " & _
                     "-d '" & jsonPayload & "' """ & apiUrl & """"""
        responseText = MacScript(macCurlCmd)
    #Else
        ' Execution path for Windows using MSXML2.ServerXMLHTTP
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        xmlHttp.Open "POST", apiUrl, False
        xmlHttp.setRequestHeader "Content-Type", "application/json"
        xmlHttp.setRequestHeader "Accept", "application/json"
        xmlHttp.send jsonPayload
        responseText = xmlHttp.responseText
    #End If
    
    ' Parse default_probability from JSON response
    Dim keyPos As Long, valStart As Long, valEnd As Long
    keyPos = InStr(responseText, """default_probability"":")
    
    If keyPos > 0 Then
        valStart = keyPos + Len("""default_probability"":")
        ' Locate ending delimiter (comma or closing brace)
        valEnd = InStr(valStart, responseText, ",")
        If valEnd = 0 Then valEnd = InStr(valStart, responseText, "}")
        
        If valEnd > valStart Then
            Dim rawVal As String
            rawVal = Trim(Mid(responseText, valStart, valEnd - valStart))
            rawVal = Replace(rawVal, ".", Application.International(xlDecimalSeparator))
            probDefault = CDbl(rawVal)
            
            ' Write probability to C14
            ws.Range("C14").Value = probDefault
            ws.Range("C14").NumberFormat = "0.00%"
            
            ' Set risk tier and recommendation
            If probDefault >= 0.5 Then
                ws.Range("C15").Value = "High Risk (Likely to Default)"
                ws.Range("C16").Value = "Decline / Reject Loan Application"
                ws.Range("C14").Font.Color = RGB(180, 0, 0)
                ws.Range("C15").Font.Color = RGB(180, 0, 0)
                ws.Range("C16").Font.Color = RGB(180, 0, 0)
            ElseIf probDefault >= 0.3 Then
                ws.Range("C15").Value = "Moderate Risk (Underwriter Review)"
                ws.Range("C16").Value = "Manual Underwriter Review Required"
                ws.Range("C14").Font.Color = RGB(200, 120, 0)
                ws.Range("C15").Font.Color = RGB(200, 120, 0)
                ws.Range("C16").Font.Color = RGB(200, 120, 0)
            Else
                ws.Range("C15").Value = "Low Risk (Unlikely to Default)"
                ws.Range("C16").Value = "Approve Loan Application"
                ws.Range("C14").Font.Color = RGB(0, 140, 40)
                ws.Range("C15").Font.Color = RGB(0, 140, 40)
                ws.Range("C16").Font.Color = RGB(0, 140, 40)
            End If
            
            ws.Range("C17").Value = "HTTP 200 OK (FastAPI)"
            ws.Range("C17").Font.Color = RGB(0, 120, 0)
            ws.Range("C18").Value = Now
            ws.Range("C18").NumberFormat = "yyyy-mm-dd hh:mm:ss"
            
            MsgBox "Loan Default Risk Assessment Completed!" & vbCrLf & vbCrLf & _
                   "Predicted Default Probability: " & Format(probDefault, "0.00%") & vbCrLf & _
                   "Risk Tier: " & ws.Range("C15").Value & vbCrLf & _
                   "Recommendation: " & ws.Range("C16").Value, _
                   vbInformation, "Bank Loan Default Prediction"
            Exit Sub
        End If
    End If
    
    ' If response could not be parsed
    ws.Range("C17").Value = "Parsing Error"
    ws.Range("C17").Font.Color = RGB(200, 0, 0)
    MsgBox "Received unexpected response from API:" & vbCrLf & vbCrLf & responseText, _
           vbExclamation, "API Response Format Error"
    Exit Sub

ErrorHandler:
    ws.Range("C17").Value = "Connection Failed"
    ws.Range("C17").Font.Color = RGB(200, 0, 0)
    MsgBox "Failed to connect to FastAPI REST endpoint at " & apiUrl & "." & vbCrLf & vbCrLf & _
           "Please ensure the FastAPI backend is running:" & vbCrLf & _
           "  uvicorn main:app --reload --port 8000" & vbCrLf & vbCrLf & _
           "Error details: " & Err.Description, _
           vbCritical, "FastAPI Service Offline"
End Sub


' -----------------------------------------------------------------------------------------
' SUBROUTINE: BatchPredictTestLoans
' PURPOSE: Iterates through each record in "BankLoan_Test Predictions" sheet, submits
'          row features to FastAPI /predict, and writes predicted default probability.
' -----------------------------------------------------------------------------------------
Public Sub BatchPredictTestLoans()
    Dim ws As Worksheet
    Dim apiUrl As String
    Dim r As Long
    Dim updatedCount As Long
    Dim jsonPayload As String
    Dim responseText As String
    
    On Error GoTo BatchErrorHandler
    
    Set ws = ActiveWorkbook.Sheets("BankLoan_Test Predictions")
    apiUrl = "http://127.0.0.1:8000/predict"
    updatedCount = 0
    
    Application.ScreenUpdating = False
    
    ' BankLoan_Test contains 10 rows located in rows 6 to 15
    For r = 6 To 15
        Dim ageVal As Double: ageVal = CDbl(ws.Cells(r, 2).Value)
        Dim employVal As Double: employVal = CDbl(ws.Cells(r, 3).Value)
        Dim addressVal As Double: addressVal = CDbl(ws.Cells(r, 4).Value)
        Dim debtincVal As Double: debtincVal = CDbl(ws.Cells(r, 5).Value)
        Dim creddebtVal As Double: creddebtVal = CDbl(ws.Cells(r, 6).Value)
        Dim othdebtVal As Double: othdebtVal = CDbl(ws.Cells(r, 7).Value)
        
        jsonPayload = "{" & _
            """AGE"":" & Replace(CStr(ageVal), ",", ".") & "," & _
            """EMPLOY"":" & Replace(CStr(employVal), ",", ".") & "," & _
            """ADDRESS"":" & Replace(CStr(addressVal), ",", ".") & "," & _
            """DEBTINC"":" & Replace(CStr(debtincVal), ",", ".") & "," & _
            """CREDDEBT"":" & Replace(CStr(creddebtVal), ",", ".") & "," & _
            """OTHDEBT"":" & Replace(CStr(othdebtVal), ",", ".") & _
            "}"
            
        #If Mac Then
            Dim macCmd As String
            macCmd = "do shell script ""curl -s -X POST -H \""Content-Type: application/json\"" " & _
                     "-d '" & jsonPayload & "' """ & apiUrl & """"""
            responseText = MacScript(macCmd)
        #Else
            Dim http As Object
            Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            http.Open "POST", apiUrl, False
            http.setRequestHeader "Content-Type", "application/json"
            http.send jsonPayload
            responseText = http.responseText
        #End If
        
        Dim keyPos As Long: keyPos = InStr(responseText, """default_probability"":")
        If keyPos > 0 Then
            Dim vStart As Long: vStart = keyPos + Len("""default_probability"":")
            Dim vEnd As Long: vEnd = InStr(vStart, responseText, ",")
            If vEnd = 0 Then vEnd = InStr(vStart, responseText, "}")
            If vEnd > vStart Then
                Dim rawNum As String
                rawNum = Trim(Mid(responseText, vStart, vEnd - vStart))
                rawNum = Replace(rawNum, ".", Application.International(xlDecimalSeparator))
                Dim prob As Double: prob = CDbl(rawNum)
                
                ws.Cells(r, 9).Value = prob
                ws.Cells(r, 9).NumberFormat = "0.00%"
                
                If prob >= 0.5 Then
                    ws.Cells(r, 10).Value = 1
                    ws.Cells(r, 11).Value = "High Risk (Default Likely)"
                    ws.Cells(r, 9).Font.Color = RGB(180, 0, 0)
                Else
                    ws.Cells(r, 10).Value = 0
                    ws.Cells(r, 11).Value = "Low Risk (Repayment Likely)"
                    ws.Cells(r, 9).Font.Color = RGB(0, 140, 40)
                End If
                updatedCount = updatedCount + 1
            End If
        End If
    Next r
    
    Application.ScreenUpdating = True
    MsgBox "Batch inference completed successfully!" & vbCrLf & _
           "Updated records: " & updatedCount & " of 10", _
           vbInformation, "Batch API Prediction Complete"
    Exit Sub

BatchErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during batch API prediction: " & Err.Description, vbCritical, "Batch Error"
End Sub


' -----------------------------------------------------------------------------------------
' SUBROUTINE: TestServerConnection
' PURPOSE: Quick ping to verify FastAPI backend is active and reachable.
' -----------------------------------------------------------------------------------------
Public Sub TestServerConnection()
    Dim apiUrl As String: apiUrl = "http://127.0.0.1:8000/health"
    Dim res As String
    
    On Error GoTo PingFail
    #If Mac Then
        res = MacScript("do shell script ""curl -s """ & apiUrl & """""")
    #Else
        Dim http As Object
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.Open "GET", apiUrl, False
        http.send
        res = http.responseText
    #End If
    
    If InStr(res, """status"":""healthy""") > 0 Then
        MsgBox "FastAPI Server is ONLINE and HEALTHY!" & vbCrLf & vbCrLf & _
               "Response: " & res, vbInformation, "Server Status: Connected"
    Else
        MsgBox "Received unexpected response from server:" & vbCrLf & res, vbExclamation, "Server Check"
    End If
    Exit Sub

PingFail:
    MsgBox "Cannot reach FastAPI server at " & apiUrl & "." & vbCrLf & _
           "Please start your server: uvicorn main:app --reload --port 8000", _
           vbCritical, "Server Status: Offline"
End Sub
