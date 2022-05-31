VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufImages 
   Caption         =   "UserForm1"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "ufImages.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'[Icons]
Public Function getAddIcon() As Variant
    Set getAddIcon = Me.icoAdd.picture
End Function

Public Function getRemoveIcon() As Variant
    Set getRemoveIcon = Me.icoRemove.picture
End Function

Public Function getRemoveRedCrossIcon() As Variant
    Set getRemoveRedCrossIcon = Me.icoRemoveWhiteBack.picture
End Function

Public Function getQuestionMarkIcon() As Variant
    Set getQuestionMarkIcon = Me.icoQuestionMark.picture
End Function

Public Function getSuccessIcon() As Variant
    Set getSuccessIcon = Me.icoSuccess.picture
End Function

Public Function getWarningIcon() As Variant
    Set getWarningIcon = Me.icoWarning.picture
End Function

Public Function getExclamationInCircle() As Variant
    Set getExclamationInCircle = Me.icoExclamationInCircle.picture
End Function

Public Function getErrorIcon() As Variant
    Set getErrorIcon = Me.icoFail.picture
End Function

Public Function getCollapseSuccessIcon() As Variant
    Set getCollapseSuccessIcon = Me.icoCollapsePass.picture
End Function

Public Function getExpandSuccessIcon() As Variant
    Set getExpandSuccessIcon = Me.icoExpandPass.picture
End Function

Public Function getCollapseWarningIcon() As Variant
    Set getCollapseWarningIcon = Me.icoCollapseWarning.picture
End Function

Public Function getExpandWarningIcon() As Variant
    Set getExpandWarningIcon = Me.icoExpandWarning.picture
End Function

Public Function getCollapseErrorIcon() As Variant
    Set getCollapseErrorIcon = Me.icoCollapseFail.picture
End Function

Public Function getExpandErrorIcon() As Variant
    Set getExpandErrorIcon = Me.icoExpandFail.picture
End Function

Public Function getOkButtonNormalImage() As Variant
    Set getOkButtonNormalImage = Me.imgOkButtonNormal.picture
End Function

Public Function getOkButtonHoverImage() As Variant
    Set getOkButtonHoverImage = Me.imgOkButtonHover.picture
End Function

Public Function getOkButtonClickImage() As Variant
    Set getOkButtonClickImage = Me.imgOkButtonClick.picture
End Function

Public Function getCancelButtonNormalImage() As Variant
    Set getCancelButtonNormalImage = Me.imgCancelButtonNormal.picture
End Function

Public Function getCancelButtonHoverImage() As Variant
    Set getCancelButtonHoverImage = Me.imgCancelButtonHover.picture
End Function

Public Function getCancelButtonClickImage() As Variant
    Set getCancelButtonClickImage = Me.imgCancelButtonClick.picture
End Function

Public Function getSelectFolderImage() As Variant
    Set getSelectFolderImage = Me.icoSelectFolder.picture
End Function

Public Function getSelectFolderWithInvalidBackImage() As Variant
    Set getSelectFolderWithInvalidBackImage = Me.icoSelectFolder_Invalid.picture
End Function

Public Function getPreviewFileErrorImage() As Variant
    Set getPreviewFileErrorImage = Me.icoPreviewFileRed.picture
End Function

Public Function getPreviewFileWarningImage() As Variant
    Set getPreviewFileWarningImage = Me.icoPreviewFileYellow.picture
End Function

Public Function getPreviewFileNormalImage() As Variant
    Set getPreviewFileNormalImage = Me.icoPreviewFileGreen.picture
End Function

Public Function getPreviewFileWhiteImage() As Variant
    Set getPreviewFileWhiteImage = Me.icoPreviewFileWhite.picture
End Function

Public Function getPreviewFileOrangeImage() As Variant
    Set getPreviewFileOrangeImage = Me.icoPreviewFileOrange.picture
End Function


Public Function getRefreshErrorImage() As Variant
    Set getRefreshErrorImage = Me.icoRefreshRed.picture
End Function

Public Function getRefreshWarningImage() As Variant
    Set getRefreshWarningImage = Me.icoRefreshYellow.picture
End Function

Public Function getRefreshNormalImage() As Variant
    Set getRefreshNormalImage = Me.icoRefreshGreen.picture
End Function



Public Function getRemoveItemErrorImage() As Variant
    Set getRemoveItemErrorImage = Me.icoRemoveItemRed.picture
End Function

Public Function getRemoveItemWarningImage() As Variant
    Set getRemoveItemWarningImage = Me.icoRemoveItemYellow.picture
End Function

Public Function getRemoveItemNormalImage() As Variant
    Set getRemoveItemNormalImage = Me.icoRemoveItemGreen.picture
End Function


Public Function getListImage() As Variant
    Set getListImage = Me.icoList.picture
End Function



'List items
Public Function getEditListItem_White() As Variant
    Set getEditListItem_White = Me.icoEditListItem_White.picture
End Function

Public Function getEditListItem_LightBlue() As Variant
    Set getEditListItem_LightBlue = Me.icoEditListItem_LightBlue.picture
End Function

Public Function getEditListItem_DarkBlue() As Variant
    Set getEditListItem_DarkBlue = Me.icoEditListItem_DarkBlue.picture
End Function

Public Function getEditListItem_Gray() As Variant
    Set getEditListItem_Gray = Me.icoEditListItem_Gray.picture
End Function


Public Function getCancelListItem_White() As Variant
    Set getCancelListItem_White = Me.icoRemoveListItem_White.picture
End Function

Public Function getCancelListItem_LightBlue() As Variant
    Set getCancelListItem_LightBlue = Me.icoRemoveListItem_LightBlue.picture
End Function

Public Function getCancelListItem_DarkBlue() As Variant
    Set getCancelListItem_DarkBlue = Me.icoRemoveListItem_DarkBlue.picture
End Function

Public Function getCancelListItem_Gray() As Variant
    Set getCancelListItem_Gray = Me.icoRemoveListItem_Gray.picture
End Function



Public Function getSortUp() As Variant
    Set getSortUp = Me.ico_SortUp.picture
End Function

Public Function getSortDown() As Variant
    Set getSortDown = Me.ico_SortDown.picture
End Function

Public Function getActiveFilter() As Variant
    Set getActiveFilter = Me.icoFilterWhite.picture
End Function

Public Function getInactiveFilter() As Variant
    Set getInactiveFilter = Me.icoFilterYellow.picture
End Function

Public Function getEllipsisBlackBack() As Variant
    Set getEllipsisBlackBack = Me.icoEllipsisBlackBack.picture
End Function

Public Function getEllipsisWhiteBack() As Variant
    Set getEllipsisWhiteBack = Me.icoEllipsisWhiteBack.picture
End Function

Public Function getInfoIcon() As Variant
    Set getInfoIcon = Me.icoInfo.picture
End Function

Public Function getExcelIcon() As Variant
    Set getExcelIcon = Me.icoExcel.picture
End Function



'[Expand/Collapse arrow icons]
Public Function getExpandArrowsIcon_Green() As Variant
    Set getExpandArrowsIcon_Green = Me.icoExpandArrowsGreen.picture
End Function

Public Function getExpandArrowsIcon_Red() As Variant
    Set getExpandArrowsIcon_Red = Me.icoExpandArrowsRed.picture
End Function

Public Function getExpandArrowsIcon_Gray() As Variant
    Set getExpandArrowsIcon_Gray = Me.icoExpandArrowsGray.picture
End Function

Public Function getExpandArrowsIcon_White() As Variant
    Set getExpandArrowsIcon_White = Me.icoExpandArrowsWhite.picture
End Function

Public Function getCollapseArrowsIcon_Green() As Variant
    Set getCollapseArrowsIcon_Green = Me.icoCollapseArrowsGreen.picture
End Function

Public Function getCollapseArrowsIcon_Red() As Variant
    Set getCollapseArrowsIcon_Red = Me.icoCollapseArrowsRed.picture
End Function

Public Function getCollapseArrowsIcon_Gray() As Variant
    Set getCollapseArrowsIcon_Gray = Me.icoCollapseArrowsGray.picture
End Function

Public Function getCollapseArrowsIcon_White() As Variant
    Set getCollapseArrowsIcon_White = Me.icoCollapseArrowsWhite.picture
End Function



'[Pointers]
Public Function getMousePointer() As Variant
    Set getMousePointer = Me.ctrlMousePointer.MouseIcon
End Function

