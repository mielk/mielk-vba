Attribute VB_Name = "modConstUi"
Option Explicit

Private Const CLASS_NAME As String = "modConstUi"
'[General] ------------------------------------------------------------------------------------------------
Public Const PIXEL_SIZE As Single = 0.75
Public Const SCROLL_BAR_WIDTH As Single = 13
Public Const PIXEL_TO_HIMETRIC As Double = 26.45833
'[Fonts & colors] -----------------------------------------------------------------------------------------
Public Const MIELK_COLOR As Long = 10907185 ' 14646353 '6066688
Public Const MIELK_COLOR_LIGHT As Long = 13017207 '16230019 '8103473
Public Const MIELK_COLOR_GRAY As Long = 13017207 '16230019 '8103473
Public Const APP_FONT_FAMILY As String = "Century Gothic"
Public Const TRANSPARENCY_LAYER_COLOR As Long = 16711679
Public Const BACKGROUND_OPACITY As Byte = 200
Public Const FULL_OPACITY As Byte = 255
'...
Public Const CONFIRM_BACK_COLOR As Long = 4496708
Public Const CONFIRM_BORDER_COLOR As Long = 3769401
Public Const CONFIRM_FONT_COLOR As Long = VBA.vbWhite
Public Const CANCEL_BACK_COLOR As Long = 2896073
Public Const CANCEL_BORDER_COLOR As Long = 2435500
Public Const CANCEL_FONT_COLOR As Long = VBA.vbWhite
Public Const NEUTRAL_BACK_COLOR As Long = 15132390
Public Const NEUTRAL_BORDER_COLOR As Long = 11382189
Public Const NEUTRAL_FONT_COLOR As Long = 3289650
'[Paddings & margins] -------------------------------------------------------------------------------------
Public Const DEFAULT_SUBFORM_OFFSET As Single = 24
'[Validation styles] --------------------------------------------------------------------------------------
Public Const VALID_BACK_COLOR As Long = 16777215
Public Const VALID_BORDER_COLOR As Long = 10921638
Public Const VALID_FONT_COLOR As Long = 4210752
Public Const INVALID_BACK_COLOR As Long = 12106214
Public Const INVALID_BORDER_COLOR As Long = 3487637
Public Const INVALID_FONT_COLOR As Long = 3487637
'[Checkboxes] ---------------------------------------------------------------------------------------------
Public Const CHECKBOX_WIDTH As Single = 10.5
'[Create control commands] --------------------------------------------------------------------------------
Public Const CREATE_BUTTON_ID As String = "Forms.CommandButton.1"
Public Const CREATE_CHECKBOX_ID As String = "Forms.CheckBox.1"
Public Const CREATE_COMBOBOX_ID As String = "Forms.ComboBox.1"
Public Const CREATE_FRAME_ID As String = "Forms.Frame.1"
Public Const CREATE_IMAGE_ID As String = "Forms.Image.1"
Public Const CREATE_LABEL_ID As String = "Forms.Label.1"
Public Const CREATE_LIST_ID As String = "Forms.ListBox.1"
Public Const CREATE_OPTION_ID As String = "Forms.OptionButton.1"
Public Const CREATE_TEXTBOX_ID As String = "Forms.TextBox.1"
Public Const CREATE_LIST_VIEW_ID As String = "MSComctlLib.ListViewCtrl.2"
'----------------------------------------------------------------------------------------------------------
