VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FValidations"
'----------------------------------------------------------------------------------------------------------

Public Function mergeValidationObjects(validations As VBA.Collection) As ValidationObject
    Dim validation As ValidationObject
    '------------------------------------------------------------------------------------------------------
    Set mergeValidationObjects = f.Create.ValidationObject(True)
    For Each validation In validations
        Call mergeValidationObjects.merge(validation)
    Next validation
End Function
