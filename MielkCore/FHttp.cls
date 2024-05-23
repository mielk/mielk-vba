VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function getResponse(url As String) As String
    Dim xmlHttp As Object       'MSXML2.ServerXMLHTTP60
    '------------------------------------------------------------------------------------------------------
    Set xmlHttp = VBA.CreateObject("MSXML2.serverXMLHTTP")
    With xmlHttp
        Call .Open("GET", url, False)
        Call .send
        getResponse = .responseText
    End With
    Set xmlHttp = Nothing
End Function
