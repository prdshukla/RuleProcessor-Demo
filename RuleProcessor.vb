Imports System.Collections.ObjectModel



Public Class RuleProcessor

    'Option Strict [OFF] is intentional to compare the objects in which specific type cannot be passed from rule manager
    'TODO: Need to revisit the code for the above case

#Region "Private Members"
    Private parens As Integer
    Private MainOp As Integer
    Private currentOp As Integer
    Private val1 As Integer
    Private val2 As Integer
    Private mainResult As Object
    Private currentResult As Object
#End Region

#Region "Public Methods"

    ''' <summary>
    ''' This function will process the rule based on the rule code
    ''' </summary>
    ''' <param name="completeRule"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ProcessRule(ByVal completeRule As String) As Boolean
        Dim expr As String
        Dim ch As Integer
        Dim pos As Integer
        Dim expr_len As Integer
        parens = 0
        MainOp = 0
        currentOp = 0
        val1 = -1
        val2 = -1
        mainResult = Nothing
        currentResult = Nothing

        Try
            If String.IsNullOrEmpty(completeRule) Then
                Throw New ArgumentNullException("completeRule", "completeRule is null")
            End If
            expr = completeRule.Replace(" ", "")
            expr_len = Len(expr)

            If expr_len = 0 Then Return False

            For pos = 0 To expr_len - 1
                ''Examine the next character.
                ch = Integer.Parse(expr.Substring(pos, 1))
                Select Case ch
                    Case 0, 1, 9
                        ProcessZeroOrOne(ch)
                    Case 2
                        ProcessTwo()
                    Case 3
                        ProcessThree()
                    Case 4
                        ProcessFour()
                    Case 5
                        ProcessFive()
                    Case 6
                        ProcessSix()
                    Case 7
                        ProcessSeven()
                End Select
            Next

            If IsNothing(mainResult) Then
                mainResult = currentResult
            End If

            MsgBox(mainResult)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' This function generates result using operand and operator codes
    ''' </summary>
    ''' <param name="value1"></param>
    ''' <param name="value2"></param>
    ''' <param name="operandCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GenerateResult(ByVal value1 As Integer, ByVal value2 As Integer, ByVal operandCode As Integer) As Boolean
        Try
            If operandCode = 6 Then
                If value1 + value2 > 0 Then
                    Return True
                Else
                    Return False
                End If

            ElseIf operandCode = 7 Then
                If value1 + value2 > 1 Then
                    Return True
                Else
                    Return False
                End If

            End If
        Catch ex As Exception
            Throw
        End Try
        Return False

    End Function

    ''' <summary>
    ''' This fn replaces the string values with equivalent numeric codes
    ''' </summary>
    ''' <param name="operatorCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NumericOperatorValue(ByVal operatorCode As String) As Integer
        Try
            Select Case UCase(operatorCode)

                Case "("
                    Return 2
                Case ")"
                    Return 3
                Case ") OR"
                    Return 4
                Case ") AND"
                    Return 5
                Case "OR"
                    Return 6
                Case "AND"
                    Return 7
                Case ""
                    Return 0
            End Select

        Catch ex As Exception
            Throw
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' Determines whether [is base data type] [the specified data type].
    ''' </summary>
    ''' <param name="dataType">Type of the data.</param>
    ''' <returns><c>true</c> if [is base data type] [the specified data type]; otherwise, <c>false</c>.</returns>
    ''' <remarks></remarks>
    Public Function IsBaseDataType(ByVal dataType As Type) As Boolean
        If dataType Is GetType(Boolean) Then
            Return True
        ElseIf dataType Is GetType(Integer) Then
            Return True
        ElseIf dataType Is GetType(Long) Then
            Return True
        ElseIf dataType Is GetType(Short) Then
            Return True
        ElseIf dataType Is GetType(String) Then
            Return True
        ElseIf dataType Is GetType(Decimal) Then
            Return True
        ElseIf dataType Is GetType(Single) Then
            Return True
        ElseIf dataType Is GetType(Double) Then
            Return True
        ElseIf dataType Is GetType(Char) Then
            Return True
        ElseIf dataType Is GetType(String) Then
            Return True
        ElseIf dataType Is GetType(Date) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Determines whether [is collection type] [the specified data type].
    ''' </summary>
    ''' <param name="dataType">Type of the data.</param>
    ''' <returns><c>true</c> if [is collection type] [the specified data type]; otherwise, <c>false</c>.</returns>
    ''' <remarks></remarks>
    Public Function IsCollectionType(ByVal dataType As Type) As Boolean
        If dataType Is Nothing Then
            Throw New ArgumentNullException("dataType")
        End If

        Return dataType.Name = GetType(Collection).Name
    End Function

    ''' <summary>
    ''' Compares the specified first data.
    ''' </summary>
    ''' <param name="firstData">The first data.</param>
    ''' <param name="secondData">The second data.</param>
    ''' <param name="objType">Type of the obj.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Compare(ByVal firstData As Object, ByVal secondData As Object, ByVal objType As Type) As Integer
        If objType Is GetType(Short) Then
            firstData = CShort(firstData)
            secondData = CShort(Val(secondData))
        ElseIf objType Is GetType(Integer) Then
            firstData = CInt(firstData)
            secondData = CInt(Val(secondData))
        ElseIf objType Is GetType(Long) Then
            firstData = CLng(firstData)
            secondData = CLng(secondData)
        ElseIf objType Is GetType(Boolean) Then
            firstData = CBool(firstData)
            secondData = CBool(secondData)
        ElseIf objType Is GetType(Single) Then
            firstData = CSng(firstData)
            secondData = CSng(Val(secondData))
        ElseIf objType Is GetType(Decimal) Then
            firstData = CDec(firstData)
            secondData = CDec(Val(secondData))
        ElseIf objType Is GetType(Double) Then
            firstData = CDbl(firstData)
            secondData = CDbl(secondData)
        ElseIf objType Is GetType(Char) Then
            firstData = CChar(firstData)
            secondData = CChar(secondData)
        ElseIf objType Is GetType(String) Then
            firstData = CStr(IIf(firstData Is Nothing, String.Empty, firstData))
            secondData = CStr(secondData)
            If secondData.ToString() = "0" Then
                If Comparer.Default.Compare(firstData, String.Empty) = 0 Or Comparer.Default.Compare(firstData, "0") = 0 Then
                    Return 0
                Else
                    Return 1
                End If
            End If
        ElseIf objType Is GetType(Date) Then
            firstData = CDate(firstData)
            secondData = CDate(secondData)
        End If

        Return Comparer.Default.Compare(firstData, secondData)
    End Function
#End Region

#Region "Private Methods"

    Private Function CheckResults(ByVal results As Boolean) As Integer
        If results = True Then
            Return 1
        Else
            Return 0
        End If
    End Function

    Private Function CheckValues(ByVal value As Integer) As Boolean
        If value = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ProcessZeroOrOne(ByVal ch As Integer)
        If val1 = -1 Then
            val1 = ch
        ElseIf val2 = -1 Then
            val2 = ch
        End If
    End Sub

    Private Sub ProcessTwo()
        parens += 1

        If Not IsNothing(currentResult) Then
            mainResult = currentResult
            currentResult = Nothing
        End If
    End Sub

    Private Sub ProcessThree()
        parens -= 1

        If val2 = -1 AndAlso Not IsNothing(currentResult) Then
            val2 = val1
            val1 = CheckResults(currentResult)
        End If

        If currentOp > 0 Then
            If Not IsNothing(currentResult) Then
                mainResult = currentResult
            End If

            currentResult = GenerateResult(val1, val2, currentOp)

        End If

        If MainOp > 0 Then
            If IsNothing(mainResult) Then
                mainResult = GenerateResult(val1, val2, MainOp)
            Else
                If val1 > -1 And IsNothing(currentResult) Then
                    mainResult = GenerateResult(val1, CheckResults(mainResult), MainOp)
                Else
                    mainResult = GenerateResult(CheckResults(mainResult), CheckResults(currentResult), MainOp)
                End If
            End If

        End If

        If currentOp = 0 AndAlso MainOp = 0 AndAlso val2 = -1 And val1 >= 0 Then
            mainResult = CheckValues(val1)
        End If
        val1 = -1
        val2 = -1
        currentOp = 0
        MainOp = 0
    End Sub

    Private Sub ProcessFour()
        parens -= 1
        If val2 = -1 AndAlso Not IsNothing(currentResult) Then
            val2 = val1
            val1 = CheckResults(currentResult)
        Else
            currentResult = CheckValues(val1)
        End If

        If currentOp > 0 Then
            currentResult = GenerateResult(val1, val2, currentOp)
            currentOp = 0
        End If

        If MainOp > 0 Then
            If IsNothing(mainResult) Then
                mainResult = GenerateResult(val1, val2, MainOp)
                MainOp = 0
            Else
                mainResult = GenerateResult(CheckResults(mainResult), CheckResults(currentResult), MainOp)
                MainOp = 0
            End If
            currentResult = Nothing
        End If

        val1 = -1
        val2 = -1
        MainOp = 6
    End Sub

    Private Sub ProcessFive()
        parens -= 1
        If val2 = -1 AndAlso Not IsNothing(currentResult) Then
            val2 = val1
            val1 = CheckResults(currentResult)
        Else
            currentResult = CheckValues(val1)
        End If

        If currentOp > 0 Then
            currentResult = GenerateResult(val1, val2, currentOp)
            currentOp = 0
        End If

        If MainOp > 0 Then
            If IsNothing(mainResult) Then
                mainResult = GenerateResult(val1, val2, MainOp)
                MainOp = 0
            Else
                mainResult = GenerateResult(CheckResults(mainResult), CheckResults(currentResult), MainOp)
                MainOp = 0
            End If
            currentResult = Nothing
        End If
        val1 = -1
        val2 = -1
        MainOp = 7
    End Sub

    Private Sub ProcessSix()
        If currentOp > 3 Then
            If Not IsNothing(currentResult) Then
                mainResult = currentResult
            End If

            currentResult = GenerateResult(val1, val2, currentOp)
            currentOp = 0
            val1 = CheckResults(currentResult)
            val2 = -1
            currentResult = Nothing
        End If

        currentOp = 6
    End Sub

    Private Sub ProcessSeven()
        If currentOp > 3 Then
            If Not IsNothing(currentResult) Then
                mainResult = currentResult
            End If

            currentResult = GenerateResult(val1, val2, currentOp)
            currentOp = 0
            val1 = CheckResults(currentResult)
            val2 = -1
            currentResult = Nothing
        End If

        currentOp = 7
    End Sub
#End Region



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ProcessRule(TextBox1.Text)
    End Sub


End Class
