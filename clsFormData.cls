VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const BOUNDRY = "7cd1d6371ec"
Dim mRequestData() As Byte
Const NL = vbCrLf

Private Sub Class_Initialize()
    'AddTextElement "method", "POST"
End Sub
Public Function GetBoundry() As String
    GetBoundry = BOUNDRY
End Function
Private Function GetContentDispositionElement(strName As String, strValue As String) As String
    GetContentDispositionElement = "--" & BOUNDRY & NL & "Content-Disposition: form-data; name=""" & strName & """" & NL & NL & strValue
End Function

Private Sub AddDataToRequest(data() As Byte)
    Dim nStart As Long, I As Long, bDataLen As Long
    
    On Error Resume Next
    nStart = UBound(mRequestData)
    On Error GoTo 0
    
    bDataLen = UBound(data)
    If nStart = 0 Then 'first
        ReDim mRequestData(0 To bDataLen)
        For I = 0 To bDataLen
            mRequestData(I) = data(I)
        Next

    Else 'rest
        ReDim Preserve mRequestData(0 To nStart + bDataLen + 1)
        For I = 0 To bDataLen
            mRequestData(nStart + 1 + I) = data(I)
        Next
    End If
End Sub
Private Function GetFileNameElement(ByVal name As String, Optional ByVal strType As String = "image/tif") As String
    GetFileNameElement = "Content-Disposition: form-data; name=""file""; filename=""" & name & """" & NL & "Content-Type: " & strType & NL & NL
End Function
Public Function AddTextElement(name As String, data As String) As Boolean
    Dim B() As Byte, strData As String
    
    If Len(name) = 0 Then
        AddTextElement = False
        Exit Function
    End If
    
    strData = GetContentDispositionElement(name, data)
    
    If Len(strData) = 0 Then
        AddTextElement = False
        Exit Function
    End If

    B = StrConv(strData, vbFromUnicode)
    AddDataToRequest B
    AddDataToRequest StrConv(NL, vbFromUnicode)
    AddTextElement = True
End Function

Public Function AddFileElement(ByVal filepath As String, ByVal strMimeType As String) As Boolean
    Dim B() As Byte, fileUploadSufix As String, fnameel As String
    Dim bFileName() As Byte, strData As String
    Dim fileUploadSufixBytes() As Byte
    
    If SafeFileExists(filepath) <> -1 Then Exit Function
    
    fileUploadSufix = "--" & BOUNDRY & NL
    fileUploadSufixBytes = StrConv(fileUploadSufix, vbFromUnicode)
    AddDataToRequest fileUploadSufixBytes
    
    fnameel = GetFileNameElement("file", strMimeType)
    bFileName = StrConv(fnameel, vbFromUnicode)
    AddDataToRequest bFileName
    
    Open filepath For Binary Access Read Lock Write As #1
    ReDim B(0 To LOF(1))
    Get #1, , B()
    AddDataToRequest B
    Close #1
    
    AddDataToRequest StrConv(NL, vbFromUnicode)
    AddFileElement = True
End Function

 Public Function GetData() As String
    AddDataToRequest StrConv("--" & BOUNDRY & "--" & NL, vbFromUnicode)
    GetData = StrConv(mRequestData, vbUnicode)
End Function
Public Function GetByteData(B() As Byte) As Boolean
    B = mRequestData
    GetByteData = True
End Function
