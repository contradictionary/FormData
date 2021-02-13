# FormData
Simple Class To prepare multipart FormData request

 ## VB6 Usage
 ```VBA
    Set xmlhttp = New WinHttpRequest
    Set formData = New clsFormData
    reqData = "{}"

    xmlhttp.Open "POST", "http://localhost/xx/xx", False
    xmlhttp.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & formData.GetBoundry()
    formData.AddTextElement "data", reqData
    strBody = formData.GetData()
    
    xmlhttp.Send strBody
    resp = xmlhttp.ResponseText
```

## VB .Net Usage 1
```VB
Dim form = New clsFormDataString()
Dim API_URL As String, bData As Byte(), data As String

API_URL = "http://localhost/XX/XX"
Dim req = WebRequest.Create(API_URL)

req.ContentType = "multipart/form-data; boundary=" & form.GetBoundry()
req.Method = "POST"

Dim reqStr = "{""XX"": ""XX"",""XX"": ""XX"",""XX"": ""Y""}"

form.AddTextElement("data", reqStr)

form.AddFileElement("D:\XX\XX\XXX.tif", "image/tif")

data = form.GetRequestString()
bData = Encoding.UTF8.GetBytes(data.ToCharArray())
req.GetRequestStream().Write(bData, 0, bData.Length)

Dim result = req.GetResponse()
Dim stream2 As Stream = result.GetResponseStream()
Dim reader2 As New StreamReader(stream2)

Dim RESULTStr = reader2.ReadToEnd()
Console.WriteLine("Response: " & vbNewLine & RESULTStr)
```

## VB .Net Usage 2
```VB
Dim form = New clsFormData()
Dim API_URL As String
API_URL = "Http://localhost/XX/XXX"

Dim req = WebRequest.Create(API_URL)

req.ContentType = "multipart/form-data; boundary=" & form.GetBoundry()
req.Method = "POST"
Dim reqStr = "{ ""data"":{ ""XXX"":1,""XXX"":0} }"

form.AddTextElement("data", reqStr)
form.AddToRequestTream(req.GetRequestStream(), 0)

Dim result = req.GetResponse()
Dim stream2 As Stream = result.GetResponseStream()
Dim reader2 As New StreamReader(stream2)

Dim RESULTStr = reader2.ReadToEnd()
Console.WriteLine("Response:" & vbNewLine & RESULTStr)
```
