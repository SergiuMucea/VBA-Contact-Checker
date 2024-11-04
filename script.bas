Sub ProcessContacts()
    Dim mainWorkbook As Workbook
    Dim resultWorkbook As Workbook
    Dim mainSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    Dim customer As String
    Dim companyCode As String
    Dim customerName As String
    Dim email As String
    Dim processedCustomers As Object
    Dim data As Variant
    Dim results() As Variant
    Dim chunkSize As Long
    Dim chunkStart As Long
    Dim chunkEnd As Long
    Dim progressForm As UserForm1
    
    ' Disable screen updating and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Create a dictionary to track processed customer numbers
    Set processedCustomers = CreateObject("Scripting.Dictionary")
    
    ' Open the main file
    Set mainWorkbook = Workbooks.Open("path\to\file.xlsx")
    Set mainSheet = mainWorkbook.Sheets(1)
    
    ' Open the result file
    Set resultWorkbook = Workbooks.Open("path\to\file.xlsm")
    Set resultSheet = resultWorkbook.Sheets(1)
    
    ' Find the last row in the main sheet
    lastRow = mainSheet.Cells(mainSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Read data into an array
    data = mainSheet.Range("A2:K" & lastRow).Value
    
    ' Initialize the result array
    ReDim results(1 To UBound(data, 1), 1 To 4)
    
    ' Initialize the result row
    resultRow = 1
    
    ' Set chunk size
    chunkSize = 10000
    
    ' Initialize and show the progress form
    Set progressForm = New UserForm1
    With progressForm
        .ProgressLabel.Caption = "Processing 0 out of " & UBound(data, 1) & " lines"
        .Show vbModeless
    End With
    
    ' Process data in chunks
    For chunkStart = 1 To UBound(data, 1) Step chunkSize
        chunkEnd = Application.Min(chunkStart + chunkSize - 1, UBound(data, 1))
        
        For i = chunkStart To chunkEnd
            customer = data(i, 1)
            
            ' Check if the customer number has already been processed
            If Not processedCustomers.exists(customer) Then
                companyCode = data(i, 2)
                customerName = data(i, 3)
                
                ' Get the email or message using the CheckContact function
                email = CheckContact(data, customer)
                
                ' Write the results to the result array
                results(resultRow, 1) = customer
                results(resultRow, 2) = companyCode
                results(resultRow, 3) = customerName
                results(resultRow, 4) = email
                
                ' Move to the next result row
                resultRow = resultRow + 1
                
                ' Mark the customer number as processed
                processedCustomers.Add customer, True
            End If
            
            ' Update the progress label
            progressForm.ProgressLabel.Caption = "Processing " & i & " out of " & UBound(data, 1) & " lines"
            DoEvents
        Next i
    Next chunkStart
    
    ' Write results to the result sheet
    resultSheet.Range("A2").Resize(resultRow - 1, 4).Value = results
    
    ' Save and close the result workbook
    resultWorkbook.Save
    
    ' Re-enable screen updating and automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Hide the progress form
    Unload progressForm
End Sub

Function CheckContact(data As Variant, customer As String) As String
    Dim email As String
    Dim foundZ5 As Boolean
    Dim i As Long
    
    foundZ5 = False
    email = ""
    
    ' Loop through each row in the array
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) = customer Then
            If data(i, 10) = "Z5" Then
                foundZ5 = True
                Exit For
            ElseIf data(i, 10) = "Z008" And data(i, 11) = "ZD" Then
                email = data(i, 8) & "*E-mail found"
            ElseIf data(i, 10) = "0002" And data(i, 11) = "Z9" Then
                email = data(i, 8) & "*E-mail found"
            End If
        End If
    Next i
    
    ' If no Z5, Z008, or 0002 found, check for any other contacts except Z2
    If email = "" And Not foundZ5 Then
        For i = LBound(data, 1) To UBound(data, 1)
            If data(i, 1) = customer Then
                If data(i, 10) <> "Z2" Then
                    email = data(i, 8) & "*E-mail found"
                    Exit For
                End If
            End If
        Next i
    End If
    
    ' If no other contacts found, check for Z2 contacts
    If email = "" And Not foundZ5 Then
        For i = LBound(data, 1) To UBound(data, 1)
            If data(i, 1) = customer Then
                If data(i, 10) = "Z2" Then
                    If data(i, 8) <> "excepted_email@email.com" Then
                        email = data(i, 8) & "*Z2 contact available"
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    
    ' If no email is found, set the result to "No email found"
    If email = "" Then
        email = "_No valid contact available" & "*No valid contact available"
    End If
    
    If foundZ5 Then
        CheckContact = "_Z5 present" & "*Z5 present"
    Else
        CheckContact = email
    End If
End Function

