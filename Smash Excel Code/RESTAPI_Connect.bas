Attribute VB_Name = "RESTAPI_Connect"
Function API_Get(apiType As String, url As String, name As String, key As String, Optional varAsync As Boolean, Optional paramDictionary As String)
'Declare variables
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    If varAsync = Null Then
        varAsync = True
    End If
'Assign parameters if needed.
    'For each item in dictionary...
        'objHTTP.setRequestHeader key, item
'Call GET procedure
    objHTTP.Open apiType, url, varAsync, name, key
    objHTTP.Send
API_Get = objHTTP.responseText
End Function
