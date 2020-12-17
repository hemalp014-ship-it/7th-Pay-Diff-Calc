Public Class સંખ્યાનુવાદક

    Dim Numbers(10) As Char

    Sub New(ByVal YourLaunguageNumbers As Char())
        Numbers = YourLaunguageNumbers
    End Sub

    Public Function ToNumber(ByVal num As Char) As Integer
        Dim value As Integer = 10
        For x = 0 To Numbers.Count - 1
            If Numbers(x) = num Then value = x
        Next
        If value = 10 Then
            If Integer.TryParse(num, 0) Then Return Integer.Parse(num)
        End If
        Return value
    End Function

    Public Function GetNumber(ByVal Number As String) As Double
        Dim ReturnValue As Double = 0.0
        Dim DotLocation As Integer = 0
        Dim setFactor As Boolean = Number.Contains(".")

        If setFactor Then
            For t = Number.Length - 1 To 0 Step -1
                If Number(t) = "." Then
                    DotLocation = t
                    Number.Replace(".", "")
                End If
            Next
        End If

        For t = Number.Length - 1 To 0 Step -1
            Dim factor As Double = 1.0
            If setFactor Then
                If t - DotLocation > 0 Then factor = Math.Pow(10, t - DotLocation - 1) Else factor = Math.Pow(10, t - DotLocation)
            Else
                factor = Math.Pow(10, t - DotLocation)
            End If
            ReturnValue += ToNumber(Number(Number.Length - 1 - t)) * factor
        Next

        Return ReturnValue
    End Function

    Public Function GetString(ByVal number As Double) As String
        Dim StringData As String = number.ToString
        Dim StringData1 As String = ""
        For i = 0 To StringData.Length - 1
            StringData1 += ToString(StringData(i))
        Next
        Return StringData1
    End Function

    Public Function GetString(ByVal StringData As String) As String
        Dim StringData1 As String = ""
        For i = 0 To StringData.Length - 1
            StringData1 += ToString(StringData(i))
        Next
        Return StringData1
    End Function

    Public Overloads Function ToString(ByVal num As Char) As String
        Select Case num
            Case "0"
                Return Numbers(0)
            Case "1"
                Return Numbers(1)
            Case "2"
                Return Numbers(2)
            Case "3"
                Return Numbers(3)
            Case "4"
                Return Numbers(4)
            Case "5"
                Return Numbers(5)
            Case "6"
                Return Numbers(6)
            Case "7"
                Return Numbers(7)
            Case "8"
                Return Numbers(8)
            Case "9"
                Return Numbers(9)
            Case Else
                Return num
        End Select

    End Function

End Class
