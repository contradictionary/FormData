Imports System.IO 'FileStream
Imports System.Text 'Encoding
Imports System.Net.WebUtility
Public Class clsFormDataString
    Private Const BOUNDRY As String = "7cd1d6371ec"
    Private mRequestData As StringBuilder
    Private Const NL = vbNewLine

    Public Sub New()
        mRequestData = New StringBuilder
        AddTextElement("method", "POST")
        'ReDim _request_data(0 To 0)
    End Sub
    Public Function GetBoundry() As String
        Return BOUNDRY
    End Function
    Private Function GetContentDispositionElement(strName As String, strValue As String) As String
        Dim prefix As String
        prefix = "--{0}{1}Content-Disposition: form-data; name=""{2}""{3}{4}"
        Return String.Format(prefix, BOUNDRY, NL, strName, NL & NL, strValue)
    End Function
    Private Sub AddDataToRequest(data As String)
        mRequestData.Append(data)
    End Sub
    Private Function GetFileNameElement(ByVal name As String, Optional ByVal strType As String = "image/tif") As String
        Dim req As String = "Content-Disposition: form-data; name=""file""; filename=""" & name & """{0}Content-Type: {1}"
        Return String.Format(req, NL, strType & NL & NL)
    End Function
    Public Function AddTextElement(name As String, data As String) As Boolean
        If name.Length = 0 Then Return False
        Dim strData As String = GetContentDispositionElement(name, data)
        If strData.Length = 0 Then Return False

        AddDataToRequest(strData)
        AddDataToRequest(NL)

        Return True
    End Function
    Public Function AddFileElement(ByVal filepath As String, ByVal strMimeType As String) As Boolean
        Dim file As FileInfo, b As String, strEnc As String
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


        Dim fileUploadSufix As String = "--" & BOUNDRY & NL

        AddDataToRequest(fileUploadSufix)


        Dim fnameel = GetFileNameElement(file.Name, strMimeType)

        AddDataToRequest(fnameel)

        b = IO.File.ReadAllText(file.FullName)
        strEnc = UrlEncode(b)
        AddDataToRequest(strEnc)
        'AddDataToRequest(b)
        AddDataToRequest(NL)

        Return True

    End Function

    Public Function GetRequestString() As String
        AddDataToRequest(NL & "--" & BOUNDRY & "--" & NL)
        Dim strReq = mRequestData.ToString()
        Console.WriteLine("Request Body:" & strReq)
        Return strReq
    End Function
End Class
