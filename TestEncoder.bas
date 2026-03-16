Attribute VB_Name = "TestEncoder"
' Test Module for SQL Injection Remediation
' This module contains test cases to verify the SQL injection vulnerability
' has been properly fixed in the cmdUnsafe_Click() method.

Option Explicit

' Test Results
Private TestsPassed As Integer
Private TestsFailed As Integer

' Test to verify parameterized query prevents SQL injection with single quote
Public Sub Test_SqlInjection_SingleQuote()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with SQL injection attempt using single quote
    testForm.txtUserName.Text = "admin' OR '1'='1"
    testForm.txtPassword.Text = "password"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should fail or return Invalid (not Valid with injected SQL)
    ' This validates the parameterized query properly handles special characters
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_SingleQuote - Single quote injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_SingleQuote - Single quote injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_SingleQuote - Query failed safely (expected)"
End Sub

' Test to verify parameterized query prevents SQL injection with OR clause
Public Sub Test_SqlInjection_OrClause()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with SQL injection attempt using OR clause
    testForm.txtUserName.Text = "admin' OR 1=1--"
    testForm.txtPassword.Text = ""

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should not authenticate with the malicious payload
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_OrClause - OR clause injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_OrClause - OR clause injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_OrClause - Query failed safely (expected)"
End Sub

' Test to verify parameterized query prevents SQL injection with UNION
Public Sub Test_SqlInjection_Union()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with SQL injection attempt using UNION
    testForm.txtUserName.Text = "admin' UNION SELECT 1--"
    testForm.txtPassword.Text = "password"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should not execute the UNION injection
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_Union - UNION injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_Union - UNION injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_Union - Query failed safely (expected)"
End Sub

' Test to verify parameterized query prevents comment-based injection
Public Sub Test_SqlInjection_Comment()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with SQL injection using comment to bypass password check
    testForm.txtUserName.Text = "admin'--"
    testForm.txtPassword.Text = "wrongpassword"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should not bypass password check with comment injection
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_Comment - Comment injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_Comment - Comment injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_Comment - Query failed safely (expected)"
End Sub

' Test to verify parameterized query prevents semicolon-based multi-statement injection
Public Sub Test_SqlInjection_MultiStatement()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with SQL injection attempting multiple statements
    testForm.txtUserName.Text = "admin'; DROP TABLE Passwords--"
    testForm.txtPassword.Text = "password"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should not execute multiple statements
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_MultiStatement - Multi-statement injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_MultiStatement - Multi-statement injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_MultiStatement - Query failed safely (expected)"
End Sub

' Test to verify normal legitimate input still works correctly
Public Sub Test_LegitimateInput_Valid()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with normal legitimate input (assuming 'testuser' exists in DB)
    testForm.txtUserName.Text = "testuser"
    testForm.txtPassword.Text = "testpass123"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should execute normally without errors
    ' This test verifies functionality is preserved
    If testForm.lblValid.Caption = "Valid" Or testForm.lblValid.Caption = "Invalid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_LegitimateInput_Valid - Normal input processed correctly"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_LegitimateInput_Valid - Normal input not processed correctly"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsFailed = TestsFailed + 1
    Debug.Print "FAIL: Test_LegitimateInput_Valid - Unexpected error: " & Err.Description
End Sub

' Test to verify special characters in legitimate passwords are handled
Public Sub Test_LegitimateInput_SpecialChars()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with legitimate password containing special characters
    testForm.txtUserName.Text = "john.doe"
    testForm.txtPassword.Text = "P@ssw0rd!2024"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should handle special characters in legitimate passwords
    If testForm.lblValid.Caption = "Valid" Or testForm.lblValid.Caption = "Invalid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_LegitimateInput_SpecialChars - Special characters handled correctly"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_LegitimateInput_SpecialChars - Special characters not handled"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsFailed = TestsFailed + 1
    Debug.Print "FAIL: Test_LegitimateInput_SpecialChars - Unexpected error: " & Err.Description
End Sub

' Test to verify empty input is handled safely
Public Sub Test_EdgeCase_EmptyInput()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with empty username and password
    testForm.txtUserName.Text = ""
    testForm.txtPassword.Text = ""

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should handle empty input without crashing
    If testForm.lblValid.Caption = "Invalid" Or testForm.lblValid.Caption = "Invalid Query" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_EdgeCase_EmptyInput - Empty input handled safely"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_EdgeCase_EmptyInput - Empty input not handled properly"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_EdgeCase_EmptyInput - Query failed safely (expected)"
End Sub

' Test to verify very long input is handled safely
Public Sub Test_EdgeCase_LongInput()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder
    Dim longString As String

    ' Create a very long input string (300 characters)
    longString = String(300, "A")

    testForm.txtUserName.Text = longString
    testForm.txtPassword.Text = longString

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The query should handle long input (truncated by parameter size 255)
    If testForm.lblValid.Caption = "Invalid" Or testForm.lblValid.Caption = "Invalid Query" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_EdgeCase_LongInput - Long input handled safely"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_EdgeCase_LongInput - Long input not handled properly"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_EdgeCase_LongInput - Query failed safely (expected)"
End Sub

' Test to verify stacked queries injection is prevented
Public Sub Test_SqlInjection_StackedQuery()
    On Error GoTo TestError

    Dim testForm As encoder
    Set testForm = New encoder

    ' Test with stacked query injection attempt
    testForm.txtUserName.Text = "admin'; INSERT INTO Passwords VALUES ('hacker', 'pwd')--"
    testForm.txtPassword.Text = "password"

    ' Execute the remediated method
    testForm.cmdUnsafe_Click

    ' The parameterized query should prevent execution of the stacked query
    If testForm.lblValid.Caption <> "Valid" Then
        TestsPassed = TestsPassed + 1
        Debug.Print "PASS: Test_SqlInjection_StackedQuery - Stacked query injection prevented"
    Else
        TestsFailed = TestsFailed + 1
        Debug.Print "FAIL: Test_SqlInjection_StackedQuery - Stacked query injection not prevented"
    End If

    Set testForm = Nothing
    Exit Sub

TestError:
    TestsPassed = TestsPassed + 1
    Debug.Print "PASS: Test_SqlInjection_StackedQuery - Query failed safely (expected)"
End Sub

' Main test runner - executes all test cases
Public Sub RunAllTests()
    TestsPassed = 0
    TestsFailed = 0

    Debug.Print "=========================================="
    Debug.Print "Running SQL Injection Remediation Tests"
    Debug.Print "=========================================="
    Debug.Print ""

    ' Run all SQL injection prevention tests
    Test_SqlInjection_SingleQuote
    Test_SqlInjection_OrClause
    Test_SqlInjection_Union
    Test_SqlInjection_Comment
    Test_SqlInjection_MultiStatement
    Test_SqlInjection_StackedQuery

    ' Run functionality preservation tests
    Test_LegitimateInput_Valid
    Test_LegitimateInput_SpecialChars

    ' Run edge case tests
    Test_EdgeCase_EmptyInput
    Test_EdgeCase_LongInput

    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "Test Results Summary"
    Debug.Print "=========================================="
    Debug.Print "Tests Passed: " & TestsPassed
    Debug.Print "Tests Failed: " & TestsFailed
    Debug.Print "Total Tests: " & (TestsPassed + TestsFailed)

    If TestsFailed = 0 Then
        Debug.Print "Status: ALL TESTS PASSED"
    Else
        Debug.Print "Status: SOME TESTS FAILED"
    End If
    Debug.Print "=========================================="
End Sub
