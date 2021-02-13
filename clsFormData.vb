Imports System.IO 'FileStream
Imports System.Text 'Encoding
Imports System.Net.Http

Public Class clsFormData
    Private Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (ByVal Destination As Long, ByVal _
    Source As Long, ByVal Length As Integer)


    Private Const BOUNDRY As String = "----------------------620036029988324162179844"
    Private _request_data As Byte()
    'Private _Stream As Stream
    Private nTotalFileLength As Long
    Private Const NL = vbNewLine

    Public Sub New()

        'ReDim _request_data(0 To 0)
    End Sub
    Private Function GetContentDispositionElement(strName As String, strValue As String) As Boolean
        Dim prefix As String
        prefix = "--{0}{1}Content-Disposition: form-data; name=""{2}""{3}{4}{5}{6}"
        Return String.Format(prefix, BOUNDRY, NL, strName, NL & NL, strValue, NL, "--" & BOUNDRY & NL)
    End Function
    Private Sub AddDataToRequest(ByVal data() As Byte)
        Dim nStart As Long = _request_data.Length
        Dim bDataLen As Long = data.Length
        If nStart = 0 Then 'first
            ReDim _request_data(0 To bDataLen)
            CopyMemory(_request_data(0), data(0), bDataLen)
        Else 'rest
            CopyMemory(_request_data(nStart - 1), data(0), bDataLen)
        End If
    End Sub
    Private Function GetFileNameElement(ByVal name As String, Optional ByVal strType As String = "image/tif") As String
        Dim req As String = "Content-Disposition: form-data; name=""file""; filename= """ & name & """{0}Content-Type: {1}"
        Return String.Format(req, NL, strType & NL & NL)
    End Function
    Public Function AddTextElement(name As String, data As String) As Boolean
        If name.Length = 0 Then Return False
        Dim strData As String = GetContentDispositionElement(name, data)
        If strData.Length = 0 Then Return False

        Dim b As Byte() = Encoding.UTF8.GetBytes(strData.ToCharArray())
        AddDataToRequest(b)

        Return True
    End Function
    Public Function AddFileElement(ByVal filepath As String, ByVal strMimeType As String) As Boolean
        Dim fileStream As FileStream = Nothing
        Dim b() As Byte, cursor As Long
        Dim file As FileInfo = Nothing
        Try
            file = New FileInfo(filepath)
        Catch ex As Exception
            'handle errors
            Debug.Assert(False)
            Return False
        End Try

        If (Not (file.Exists())) Then
            Return False
        End If

        nTotalFileLength = nTotalFileLength
        Dim bFileName As Byte() = Encoding.UTF8.GetBytes(GetFileNameElement(file.Name).ToCharArray())
        AddDataToRequest(bFileName)

        'Dim br As BinaryReader = New BinaryReader(New FileStream(file.FullName, FileMode.Open, FileAccess.Read), UTF8Encoding.UTF8)
        'br.

        'Dim fstream As FileStream = New FileStream(file.FullName, FileMode.Open, FileAccess.Read)
        'ReDim b(0 To file.Length - 1)
        'cursor = 0
        'While fstream.CanRead
        '    b(cursor) = fstream.ReadByte()
        'End While

        b = System.IO.File.ReadAllBytes(file.FullName)

        Dim fileUploadSufix As String = NL & "--" & BOUNDRY & NL ' "--" & vbNewLine
        Dim fileUploadSufixBytes As Byte() = Encoding.UTF8.GetBytes(fileUploadSufix.ToCharArray())
        AddDataToRequest(fileUploadSufixBytes)
        AddDataToRequest(b)

        nTotalFileLength = nTotalFileLength + file.Length
        Return True
    End Function

    Public Function AddToRequestTream(stram As Stream) As Boolean

        Return True
    End Function
    Public Function GetByteData(b() As Byte) As Boolean
        b = _request_data
        Return True
    End Function
End Class
