VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FCast"
'----------------------------------------------------------------------------------------------------------





Public Function toProp(value As Variant) As UIProp
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is UIProp Then
                Set toProp = value
            End If
        End If
    End If
End Function

Public Function toMsFormsControl(value As Variant) As MSForms.control
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is MSForms.control Then
                Set toMsFormsControl = value
            End If
        End If
    End If
End Function

Public Function toMsFormsFrame(value As Variant) As MSForms.frame
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is MSForms.frame Then
                Set toMsFormsFrame = value
            End If
        End If
    End If
End Function

Public Function toUiIcon(value As Variant) As UIIcon
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is UIIcon Then
                Set toUiIcon = value
            End If
        End If
    End If
End Function
    
Public Function toUiLabel(value As Variant) As UILabel
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is UILabel Then
                Set toUiLabel = value
            End If
        End If
    End If
End Function

Public Function toUiButton(value As Variant) As UIButton
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is UIButton Then
                Set toUiButton = value
            End If
        End If
    End If
End Function

Public Function toIControl(value As Variant) As IControl
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is IControl Then
                Set toIControl = value
            End If
        End If
    End If
End Function

Public Function toIPanel(value As Variant) As IPanel
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is IPanel Then
                Set toIPanel = value
            End If
        End If
    End If
End Function

