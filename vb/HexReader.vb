Imports System.IO
Class HexReader
    Dim filename As String = ""
    Public stored As String = ""
    Public skip As Integer = 0
    Dim br As BinaryReader
    Sub New(path As String)
        filename = path
        br = New BinaryReader(File.Open(filename, FileMode.Open))
    End Sub

    Sub jump()
        stored = ""
        Do While skip > 0
            ReadByte()
        Loop
        'stored = ""
    End Sub
    Function Read7Bit() As Long
        stored = ""
        Dim result As Long = 0
        Dim temp As Byte = ReadByte()
        result = temp Mod 128
        While temp \ 128 = 1
            temp = ReadByte()
            If result < 4294967167 Then
                result = result * 128 + temp Mod 128
            End If
        End While
        Return result
    End Function
    Function ReadSIByte(n As Integer) As Integer
        stored = ""
        Dim result As Integer = 0
        Dim temp As Byte
        For i = 1 To n
            temp = ReadByte()
            result *= 256
            result += temp
        Next
        Return result
    End Function

    Function eof() As Boolean
        Return br.BaseStream.Position >= br.BaseStream.Length
    End Function
    Function ReadByte(Optional clear As Boolean = False) As Byte
        If clear Then
            stored = ""
        End If
        Dim res As Byte = br.ReadByte
        If skip > 0 Then
            skip -= 1
        End If
        stored &= res.ToString("X2")
        Return res
    End Function
    Function ReadSByte() As String
        Return ReadByte.ToString("X2")
    End Function
    Function ReadSByte(n As Integer) As String
        stored = ""
        Dim r As String = ""
        For i = 1 To n
            r &= ReadByte.ToString("X2")
        Next
        Return r
    End Function

    Function ReadByteCom() As SByte
        stored = ""
        Dim temp As Byte = ReadByte()
        If temp <= 127 Then
            Return temp
        Else
            Return temp - 256
        End If
    End Function


    Function ReadTByte() As String
        stored = ""
        Return Chr(ReadByte)
    End Function
    Function ReadTByte(n As Integer) As String
        stored = ""
        Dim r As String = ""
        For i = 1 To n
            r &= Chr(ReadByte)
        Next
        Return r
    End Function

    Sub Close()
        br.Dispose()
    End Sub
End Class
