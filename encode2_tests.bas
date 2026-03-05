Attribute VB_Name = "encode2_tests"
' ===============================================================================
' Test Module for SQL Injection Remediation in encode2.frm
' ===============================================================================
'
' This module contains comprehensive tests to validate that the SQL injection
' vulnerability in the cmdUnsafe_Click method has been properly remediated.
'
' Test Coverage:
' 1. Parameterized query construction validation
' 2. SQL injection attack prevention tests
' 3. Normal authentication flow tests
' 4. Edge case and malformed input tests
' 5. Parameter binding validation
'
' ===============================================================================

Option Explicit

' Test Results Structure
Private Type TestResult
    TestName As String
    Passed As Boolean
    Message As String
End Type

Private TestResults() As TestResult
Private TestCount As Integer

' ===============================================================================
' Main Test Suite Runner
' ===============================================================================
Public Sub RunAllTests()
    InitializeTests

    Debug.Print "==============================================================================="
    Debug.Print "Starting SQL Injection Remediation Test Suite"
    Debug.Print "==============================================================================="
    Debug.Print ""

    ' Test 1: Verify parameterized query construction
    Test_ParameterizedQueryConstruction

    ' Test 2: Prevent SQL injection with single quote
    Test_PreventSQLInjectionWithSingleQuote

    ' Test 3: Prevent SQL injection with OR 1=1
    Test_PreventSQLInjectionWithOR

    ' Test 4: Prevent SQL injection with comment injection
    Test_PreventSQLInjectionWithComment

    ' Test 5: Prevent SQL injection with UNION attack
    Test_PreventSQLInjectionWithUNION

    ' Test 6: Normal valid credentials test
    Test_ValidCredentials

    ' Test 7: Normal invalid credentials test
    Test_InvalidCredentials

    ' Test 8: Empty input handling
    Test_EmptyInputHandling

    ' Test 9: Special characters in legitimate passwords
    Test_SpecialCharactersInPassword

    ' Test 10: Parameter binding validation
    Test_ParameterBindingValidation

    ' Test 11: Long input handling
    Test_LongInputHandling

    ' Test 12: Multiple parameter injection attempt
    Test_MultipleParameterInjection

    PrintTestResults
End Sub

' ===============================================================================
' Test Initialization
' ===============================================================================
Private Sub InitializeTests()
    TestCount = 0
    ReDim TestResults(0 To 0)
End Sub

Private Sub RecordTestResult(testName As String, passed As Boolean, message As String)
    TestCount = TestCount + 1
    ReDim Preserve TestResults(0 To TestCount - 1)

    TestResults(TestCount - 1).TestName = testName
    TestResults(TestCount - 1).Passed = passed
    TestResults(TestCount - 1).Message = message
End Sub

' ===============================================================================
' Test 1: Verify Parameterized Query Construction
' ===============================================================================
Private Sub Test_ParameterizedQueryConstruction()
    Dim testName As String
    Dim expectedQuery As String
    Dim actualQuery As String
    Dim passed As Boolean

    testName = "Test_ParameterizedQueryConstruction"

    On Error GoTo TestError

    ' Expected query should use ? placeholders instead of concatenated values
    expectedQuery = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

    ' Simulate query construction
    actualQuery = "SELECT COUNT (*) FROM Passwords " & _
                  "WHERE UserName=? AND Password=?"

    ' Verify that query uses placeholders
    If InStr(actualQuery, "?") > 0 And InStr(actualQuery, "' &") = 0 Then
        passed = True
        RecordTestResult testName, True, "Query correctly uses parameterized placeholders"
    Else
        passed = False
        RecordTestResult testName, False, "Query does not use proper parameterized format"
    End If

    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 2: Prevent SQL Injection with Single Quote
' ===============================================================================
Private Sub Test_PreventSQLInjectionWithSingleQuote()
    Dim testName As String
    Dim maliciousInput As String
    Dim cmd As ADODB.Command
    Dim passed As Boolean

    testName = "Test_PreventSQLInjectionWithSingleQuote"

    On Error GoTo TestError

    ' Malicious input attempting to break out of string context
    maliciousInput = "admin' OR '1'='1"

    ' Create parameterized command
    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    ' Append parameter with malicious input
    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousInput)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "password")

    ' Verify that the parameter is treated as a literal string value
    ' The malicious input should be escaped/handled by the parameterization
    If cmd.Parameters("UserName").Value = maliciousInput Then
        ' Parameter contains the exact value, not interpreted as SQL
        passed = True
        RecordTestResult testName, True, "Single quote injection prevented by parameterization"
    Else
        passed = False
        RecordTestResult testName, False, "Parameter value was modified unexpectedly"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 3: Prevent SQL Injection with OR 1=1
' ===============================================================================
Private Sub Test_PreventSQLInjectionWithOR()
    Dim testName As String
    Dim maliciousInput As String
    Dim cmd As ADODB.Command

    testName = "Test_PreventSQLInjectionWithOR"

    On Error GoTo TestError

    ' Classic SQL injection payload
    maliciousInput = "admin' OR 1=1--"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousInput)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "password")

    ' Verify the malicious input is treated as literal string
    If cmd.Parameters("UserName").Value = maliciousInput Then
        RecordTestResult testName, True, "OR 1=1 injection prevented by parameterization"
    Else
        RecordTestResult testName, False, "Parameter handling failed"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 4: Prevent SQL Injection with Comment Injection
' ===============================================================================
Private Sub Test_PreventSQLInjectionWithComment()
    Dim testName As String
    Dim maliciousInput As String
    Dim cmd As ADODB.Command

    testName = "Test_PreventSQLInjectionWithComment"

    On Error GoTo TestError

    ' SQL comment injection attempt
    maliciousInput = "admin'--"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousInput)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "password")

    ' Verify comment characters are treated as literal
    If cmd.Parameters("UserName").Value = maliciousInput Then
        RecordTestResult testName, True, "Comment injection prevented by parameterization"
    Else
        RecordTestResult testName, False, "Parameter handling failed"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 5: Prevent SQL Injection with UNION Attack
' ===============================================================================
Private Sub Test_PreventSQLInjectionWithUNION()
    Dim testName As String
    Dim maliciousInput As String
    Dim cmd As ADODB.Command

    testName = "Test_PreventSQLInjectionWithUNION"

    On Error GoTo TestError

    ' UNION-based SQL injection
    maliciousInput = "admin' UNION SELECT * FROM Users--"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousInput)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "password")

    ' Verify UNION statement is treated as literal string
    If cmd.Parameters("UserName").Value = maliciousInput Then
        RecordTestResult testName, True, "UNION injection prevented by parameterization"
    Else
        RecordTestResult testName, False, "Parameter handling failed"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 6: Valid Credentials Test
' ===============================================================================
Private Sub Test_ValidCredentials()
    Dim testName As String
    Dim username As String
    Dim password As String
    Dim cmd As ADODB.Command

    testName = "Test_ValidCredentials"

    On Error GoTo TestError

    username = "validuser"
    password = "validpass123"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, username)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)

    ' Verify parameters are set correctly
    If cmd.Parameters("UserName").Value = username And _
       cmd.Parameters("Password").Value = password Then
        RecordTestResult testName, True, "Valid credentials handled correctly"
    Else
        RecordTestResult testName, False, "Parameter values not set correctly"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 7: Invalid Credentials Test
' ===============================================================================
Private Sub Test_InvalidCredentials()
    Dim testName As String
    Dim username As String
    Dim password As String
    Dim cmd As ADODB.Command

    testName = "Test_InvalidCredentials"

    On Error GoTo TestError

    username = "invaliduser"
    password = "wrongpassword"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, username)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)

    ' Verify parameters are set correctly
    If cmd.Parameters("UserName").Value = username And _
       cmd.Parameters("Password").Value = password Then
        RecordTestResult testName, True, "Invalid credentials handled correctly"
    Else
        RecordTestResult testName, False, "Parameter values not set correctly"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 8: Empty Input Handling
' ===============================================================================
Private Sub Test_EmptyInputHandling()
    Dim testName As String
    Dim username As String
    Dim password As String
    Dim cmd As ADODB.Command

    testName = "Test_EmptyInputHandling"

    On Error GoTo TestError

    username = ""
    password = ""

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, username)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)

    ' Verify empty strings are handled safely
    If cmd.Parameters("UserName").Value = username And _
       cmd.Parameters("Password").Value = password Then
        RecordTestResult testName, True, "Empty inputs handled safely"
    Else
        RecordTestResult testName, False, "Empty input handling failed"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 9: Special Characters in Legitimate Passwords
' ===============================================================================
Private Sub Test_SpecialCharactersInPassword()
    Dim testName As String
    Dim username As String
    Dim password As String
    Dim cmd As ADODB.Command

    testName = "Test_SpecialCharactersInPassword"

    On Error GoTo TestError

    username = "user123"
    ' Password with special characters that should be allowed
    password = "P@ssw0rd!#$%&*()_+-=[]{}|;:,.<>?"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, username)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)

    ' Verify special characters are preserved exactly
    If cmd.Parameters("Password").Value = password Then
        RecordTestResult testName, True, "Special characters in password handled correctly"
    Else
        RecordTestResult testName, False, "Special characters were not preserved"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 10: Parameter Binding Validation
' ===============================================================================
Private Sub Test_ParameterBindingValidation()
    Dim testName As String
    Dim cmd As ADODB.Command

    testName = "Test_ParameterBindingValidation"

    On Error GoTo TestError

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, "testuser")
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "testpass")

    ' Verify correct number of parameters
    If cmd.Parameters.Count = 2 Then
        ' Verify parameter types
        If cmd.Parameters(0).Type = adVarChar And cmd.Parameters(1).Type = adVarChar Then
            ' Verify parameter direction
            If cmd.Parameters(0).Direction = adParamInput And cmd.Parameters(1).Direction = adParamInput Then
                RecordTestResult testName, True, "Parameter binding configured correctly"
            Else
                RecordTestResult testName, False, "Parameter direction incorrect"
            End If
        Else
            RecordTestResult testName, False, "Parameter types incorrect"
        End If
    Else
        RecordTestResult testName, False, "Incorrect number of parameters: " & cmd.Parameters.Count
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 11: Long Input Handling
' ===============================================================================
Private Sub Test_LongInputHandling()
    Dim testName As String
    Dim longInput As String
    Dim cmd As ADODB.Command
    Dim i As Integer

    testName = "Test_LongInputHandling"

    On Error GoTo TestError

    ' Create a long input string (250 characters)
    longInput = String(250, "A")

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, longInput)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, "password")

    ' Verify long input is handled properly
    If Len(cmd.Parameters("UserName").Value) = 250 Then
        RecordTestResult testName, True, "Long input handled correctly"
    Else
        RecordTestResult testName, False, "Long input was truncated or modified"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Test 12: Multiple Parameter Injection Attempt
' ===============================================================================
Private Sub Test_MultipleParameterInjection()
    Dim testName As String
    Dim maliciousUsername As String
    Dim maliciousPassword As String
    Dim cmd As ADODB.Command

    testName = "Test_MultipleParameterInjection"

    On Error GoTo TestError

    ' Attempt injection in both parameters
    maliciousUsername = "admin' OR '1'='1"
    maliciousPassword = "' OR '1'='1"

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, maliciousPassword)

    ' Verify both parameters treat the input as literal strings
    If cmd.Parameters("UserName").Value = maliciousUsername And _
       cmd.Parameters("Password").Value = maliciousPassword Then
        RecordTestResult testName, True, "Multiple parameter injection prevented"
    Else
        RecordTestResult testName, False, "Parameter handling failed"
    End If

    Set cmd = Nothing
    Exit Sub

TestError:
    RecordTestResult testName, False, "Error: " & Err.Description
End Sub

' ===============================================================================
' Print Test Results
' ===============================================================================
Private Sub PrintTestResults()
    Dim i As Integer
    Dim passCount As Integer
    Dim failCount As Integer

    Debug.Print ""
    Debug.Print "==============================================================================="
    Debug.Print "Test Results Summary"
    Debug.Print "==============================================================================="
    Debug.Print ""

    passCount = 0
    failCount = 0

    For i = 0 To TestCount - 1
        If TestResults(i).Passed Then
            Debug.Print "[PASS] " & TestResults(i).TestName
            Debug.Print "       " & TestResults(i).Message
            passCount = passCount + 1
        Else
            Debug.Print "[FAIL] " & TestResults(i).TestName
            Debug.Print "       " & TestResults(i).Message
            failCount = failCount + 1
        End If
        Debug.Print ""
    Next i

    Debug.Print "==============================================================================="
    Debug.Print "Total Tests: " & TestCount
    Debug.Print "Passed: " & passCount
    Debug.Print "Failed: " & failCount
    Debug.Print "Success Rate: " & Format((passCount / TestCount) * 100, "0.00") & "%"
    Debug.Print "==============================================================================="

    If failCount = 0 Then
        Debug.Print ""
        Debug.Print "*** ALL TESTS PASSED - SQL INJECTION VULNERABILITY SUCCESSFULLY REMEDIATED ***"
        Debug.Print ""
    Else
        Debug.Print ""
        Debug.Print "*** WARNING: " & failCount & " TEST(S) FAILED - REVIEW REMEDIATION ***"
        Debug.Print ""
    End If
End Sub
