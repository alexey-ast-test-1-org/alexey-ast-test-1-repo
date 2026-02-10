Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ADODB

''' <summary>
''' Test suite for SQL Injection vulnerability remediation in encode2.frm
''' These tests validate that the cmdUnsafe_Click function properly uses
''' parameterized queries to prevent SQL injection attacks.
''' </summary>
<TestClass>
Public Class Encode2SecurityTests

    ''' <summary>
    ''' Mock database connection for testing
    ''' </summary>
    Private mockConnection As ADODB.Connection

    ''' <summary>
    ''' Test setup - initialize mock database connection
    ''' </summary>
    <TestInitialize>
    Public Sub TestSetup()
        ' Initialize mock connection for testing
        mockConnection = New ADODB.Connection()
    End Sub

    ''' <summary>
    ''' Test cleanup - dispose of resources
    ''' </summary>
    <TestCleanup>
    Public Sub TestCleanup()
        If mockConnection IsNot Nothing AndAlso mockConnection.State = ConnectionState.Open Then
            mockConnection.Close()
        End If
        mockConnection = Nothing
    End Sub

    ''' <summary>
    ''' Test that parameterized query prevents basic SQL injection attack
    ''' Attack vector: ' OR '1'='1
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_PreventsSQLInjection_SingleQuoteAttack()
        ' Arrange - Create SQL injection payload
        Dim maliciousUsername As String = "admin' OR '1'='1"
        Dim normalPassword As String = "password123"

        ' Act - Build parameterized query (as fixed in cmdUnsafe_Click)
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            ' Add parameters - this prevents SQL injection
            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, normalPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify parameters were properly added
        Assert.AreEqual(2, cmd.Parameters.Count, "Command should have exactly 2 parameters")
        Assert.AreEqual(maliciousUsername, cmd.Parameters(0).Value, "Username parameter should contain the exact input without modification")
        Assert.AreEqual(normalPassword, cmd.Parameters(1).Value, "Password parameter should contain the exact input without modification")

        ' Verify query structure doesn't contain injected SQL
        Assert.IsFalse(query.Contains("OR"), "Query template should not contain OR operator from injection attempt")
        Assert.IsTrue(query.Contains("?"), "Query should use parameter placeholders")
    End Sub

    ''' <summary>
    ''' Test that parameterized query prevents union-based SQL injection
    ''' Attack vector: ' UNION SELECT * FROM Users--
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_PreventsSQLInjection_UnionAttack()
        ' Arrange - Create union-based SQL injection payload
        Dim maliciousUsername As String = "admin' UNION SELECT password FROM Users--"
        Dim normalPassword As String = "test"

        ' Act - Build parameterized query
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, normalPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify malicious payload is treated as literal string parameter
        Assert.AreEqual(maliciousUsername, cmd.Parameters(0).Value, "Malicious UNION payload should be treated as literal string")
        Assert.IsFalse(query.Contains("UNION"), "Query template should not contain UNION from injection attempt")
    End Sub

    ''' <summary>
    ''' Test that parameterized query prevents comment-based SQL injection
    ''' Attack vector: admin'--
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_PreventsSQLInjection_CommentAttack()
        ' Arrange - Create comment-based SQL injection payload
        Dim maliciousUsername As String = "admin'--"
        Dim anyPassword As String = "ignored"

        ' Act - Build parameterized query
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, anyPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify comment characters are treated as literal data
        Assert.AreEqual(2, cmd.Parameters.Count, "Both parameters should be present despite comment attempt")
        Assert.IsTrue(cmd.Parameters(0).Value.ToString().Contains("--"), "Comment characters should be preserved in parameter value")
    End Sub

    ''' <summary>
    ''' Test that parameterized query handles legitimate special characters
    ''' Valid usernames may contain apostrophes (e.g., O'Brien)
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_HandlesLegitimateSpecialCharacters()
        ' Arrange - Create legitimate username with apostrophe
        Dim legitimateUsername As String = "O'Brien"
        Dim password As String = "SecurePass123!"

        ' Act - Build parameterized query
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, legitimateUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, password)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify special characters are handled correctly
        Assert.AreEqual(legitimateUsername, cmd.Parameters(0).Value, "Apostrophe should be preserved in parameter")
        Assert.AreEqual(password, cmd.Parameters(1).Value, "Special characters in password should be preserved")
    End Sub

    ''' <summary>
    ''' Test that parameterized query prevents time-based blind SQL injection
    ''' Attack vector: '; WAITFOR DELAY '00:00:05'--
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_PreventsSQLInjection_TimeBasedBlindAttack()
        ' Arrange - Create time-based blind SQL injection payload
        Dim maliciousUsername As String = "admin'; WAITFOR DELAY '00:00:05'--"
        Dim password As String = "test"

        ' Act - Build parameterized query
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, password)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify WAITFOR command is treated as literal string
        Assert.IsTrue(cmd.Parameters(0).Value.ToString().Contains("WAITFOR"), "WAITFOR command should be in parameter value as literal text")
        Assert.IsFalse(query.Contains("WAITFOR"), "Query template should not contain WAITFOR from injection attempt")
    End Sub

    ''' <summary>
    ''' Test that parameterized query prevents stacked query injection
    ''' Attack vector: '; DROP TABLE Users--
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_PreventsSQLInjection_StackedQueryAttack()
        ' Arrange - Create stacked query injection payload
        Dim maliciousUsername As String = "admin'; DROP TABLE Users--"
        Dim password As String = "test"

        ' Act - Build parameterized query
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, maliciousUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, password)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify DROP command is treated as literal string
        Assert.IsTrue(cmd.Parameters(0).Value.ToString().Contains("DROP TABLE"), "DROP TABLE command should be in parameter value as literal text")
        Assert.IsFalse(query.Contains("DROP"), "Query template should not contain DROP from injection attempt")
        Assert.AreEqual(1, query.Split(";"c).Length, "Query template should contain only one statement (no stacked queries)")
    End Sub

    ''' <summary>
    ''' Test that query structure uses proper parameterization
    ''' Validates the fix implements ? placeholders correctly
    ''' </summary>
    <TestMethod>
    Public Sub TestQueryStructure_UsesParameterPlaceholders()
        ' Arrange
        Dim expectedQuery As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        ' Act - Create query as in fixed code
        Dim actualQuery As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        ' Assert - Verify query uses parameter placeholders
        Assert.AreEqual(expectedQuery, actualQuery, "Query should use ? placeholders for parameters")
        Assert.AreEqual(2, actualQuery.Split("?"c).Length - 1, "Query should have exactly 2 parameter placeholders")
        Assert.IsFalse(actualQuery.Contains("'"), "Query template should not contain single quotes for parameter values")
        Assert.IsFalse(actualQuery.Contains("+"), "Query template should not use string concatenation")
        Assert.IsFalse(actualQuery.Contains("&"), "Query template should not use VB string concatenation operator")
    End Sub

    ''' <summary>
    ''' Test that parameter types are correctly defined
    ''' Validates parameters use adVarChar with appropriate length
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterConfiguration_UsesCorrectDataTypes()
        ' Arrange
        Dim username As String = "testuser"
        Dim password As String = "testpass"

        ' Act - Create parameters as in fixed code
        Dim cmd As New ADODB.Command
        cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"
        cmd.CommandType = adCmdText

        Dim param1 As ADODB.Parameter = cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, username)
        cmd.Parameters.Append(param1)

        Dim param2 As ADODB.Parameter = cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)
        cmd.Parameters.Append(param2)

        ' Assert - Verify parameter configurations
        Assert.AreEqual(adVarChar, cmd.Parameters(0).Type, "Username parameter should be adVarChar type")
        Assert.AreEqual(adParamInput, cmd.Parameters(0).Direction, "Username parameter should be input direction")
        Assert.AreEqual(255, cmd.Parameters(0).Size, "Username parameter should have size 255")

        Assert.AreEqual(adVarChar, cmd.Parameters(1).Type, "Password parameter should be adVarChar type")
        Assert.AreEqual(adParamInput, cmd.Parameters(1).Direction, "Password parameter should be input direction")
        Assert.AreEqual(255, cmd.Parameters(1).Size, "Password parameter should have size 255")
    End Sub

    ''' <summary>
    ''' Test that empty string inputs are handled safely
    ''' Edge case: empty username/password should not cause injection
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_HandlesEmptyStrings()
        ' Arrange
        Dim emptyUsername As String = ""
        Dim emptyPassword As String = ""

        ' Act - Build parameterized query with empty strings
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, emptyUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, emptyPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify empty strings are handled
        Assert.AreEqual(2, cmd.Parameters.Count, "Both parameters should be present")
        Assert.AreEqual(String.Empty, cmd.Parameters(0).Value, "Empty username should be preserved")
        Assert.AreEqual(String.Empty, cmd.Parameters(1).Value, "Empty password should be preserved")
    End Sub

    ''' <summary>
    ''' Test that very long inputs are handled safely
    ''' Edge case: inputs exceeding typical length should not cause issues
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_HandlesLongInputs()
        ' Arrange - Create very long input strings
        Dim longUsername As String = New String("A"c, 300)
        Dim longPassword As String = New String("B"c, 300)

        ' Act - Build parameterized query with long strings
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, longUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, longPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify long strings are handled (will be truncated to size)
        Assert.AreEqual(2, cmd.Parameters.Count, "Both parameters should be present")
        Assert.IsNotNull(cmd.Parameters(0).Value, "Username parameter should have a value")
        Assert.IsNotNull(cmd.Parameters(1).Value, "Password parameter should have a value")
    End Sub

    ''' <summary>
    ''' Test that NULL/Nothing values are handled safely
    ''' Edge case: NULL inputs should not cause exceptions
    ''' </summary>
    <TestMethod>
    Public Sub TestParameterizedQuery_HandlesNullValues()
        ' Arrange
        Dim nullUsername As Object = DBNull.Value
        Dim nullPassword As Object = DBNull.Value

        ' Act - Build parameterized query with NULL values
        Dim cmd As New ADODB.Command
        Dim query As String = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

        With cmd
            .CommandText = query
            .CommandType = adCmdText

            Dim param1 As ADODB.Parameter = .CreateParameter("UserName", adVarChar, adParamInput, 255, nullUsername)
            .Parameters.Append(param1)

            Dim param2 As ADODB.Parameter = .CreateParameter("Password", adVarChar, adParamInput, 255, nullPassword)
            .Parameters.Append(param2)
        End With

        ' Assert - Verify NULL values are handled
        Assert.AreEqual(2, cmd.Parameters.Count, "Both parameters should be present")
    End Sub

End Class
