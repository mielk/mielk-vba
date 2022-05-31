VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIPropsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "UIPropsEnum"
'----------------------------------------------------------------------------------------------------------
'[Size]
Private Const PROP_MIN_WIDTH As String = "minWidth"
Private Const PROP_WIDTH As String = "width"
Private Const PROP_MAX_WIDTH As String = "maxWidth"
Private Const PROP_MIN_HEIGHT As String = "minHeight"
Private Const PROP_HEIGHT As String = "height"
Private Const PROP_MAX_HEIGHT As String = "maxHeight"
'[Position]
Private Const PROP_TOP As String = "top"
Private Const PROP_RIGHT As String = "right"
Private Const PROP_BOTTOM As String = "bottom"
Private Const PROP_LEFT As String = "left"
'[Margins & paddings]
Private Const PROP_MARGIN As String = "margin"
Private Const PROP_MARGIN_VERTICAL As String = "marginVertical"
Private Const PROP_MARGIN_HORIZONTAL As String = "marginHorizontal"
Private Const PROP_MARGIN_TOP As String = "marginTop"
Private Const PROP_MARGIN_RIGHT As String = "marginRight"
Private Const PROP_MARGIN_BOTTOM As String = "marginBottom"
Private Const PROP_MARGIN_LEFT As String = "marginLeft"
Private Const PROP_PADDING As String = "padding"
Private Const PROP_PADDING_VERTICAL As String = "paddingVertical"
Private Const PROP_PADDING_HORIZONTAL As String = "paddingHorizontal"
Private Const PROP_PADDING_TOP As String = "paddingTop"
Private Const PROP_PADDING_RIGHT As String = "paddingRight"
Private Const PROP_PADDING_BOTTOM As String = "paddingBottom"
Private Const PROP_PADDING_LEFT As String = "paddingLeft"
'[Background]
Private Const PROP_BACKGROUND_STYLE As String = "backgroundStyle"
Private Const PROP_BACKGROUND_COLOR As String = "backgroundColor"
'[Font]
Private Const PROP_FONT_SIZE As String = "fontSize"
Private Const PROP_FONT_COLOR As String = "fontColor"
Private Const PROP_FONT_FAMILY As String = "fontFamily"
Private Const PROP_FONT_BOLD As String = "fontBold"
Private Const PROP_FONT_ITALIC As String = "fontItalic"
'[Border]
Private Const PROP_BORDERS As String = "borders"
Private Const PROP_BORDER_VISIBLE As String = "borderVisible"
Private Const PROP_BORDER_COLOR As String = "borderColor"
Private Const PROP_BORDER_WIDTH As String = "borderWidth"
Private Const PROP_BORDER_STYLE As String = "borderStyle"
'[Text]
Private Const PROP_HORIZONTAL_ALIGNMENT As String = "horizontalAlignment"
Private Const PROP_VERTICAL_ALIGNMENT As String = "verticalAlignment"
Private Const PROP_TEXT_UPPERCASE As String = "textUppercase"
'[Pictures]
Private Const PROP_PICTURE As String = "picture"
Private Const PROP_PICTURE_WIDTH As String = "pictureWidth"
Private Const PROP_PICTURE_HEIGHT As String = "pictureHeight"
Private Const PROP_PICTURE_MARGIN As String = "pictureMargin"

'[CONTROL SPECIFIC]
'[Combobox]
Private Const PROP_MATCH_REQUIRED As String = "matchRequired"
'[Checkbox]
Private Const PROP_TRIPLE_STATE As String = "tripleState"
'[List]
Private Const PROP_LINE_HEIGHT As String = "lineHeight"
Private Const PROP_SCROLL_BARS As String = "scrollBars"
'[Textbox]
'allow-only-numbers
'allow-negative-numbers
'is-multiline
'is-password
'number-format

'[Worksheet]
Private Const PROP_INDENT_LEVEL As String = "indent"
Private Const PROP_WRAP_TEXT As String = "wrapText"
Private Const PROP_AUTO_FIT As String = "autoFit"
Private Const PROP_NUMBER_FORMAT As String = "numberFormat"

'----------------------------------------------------------------------------------------------------------
Private pDefaultPropsArray() As Variant
Private pEmptyPropsArray() As Variant
Private pPropsByName As Scripting.Dictionary
Private pPropsByAliases As Scripting.Dictionary
Private pPropsById As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------




Private Sub Class_Initialize()
    Call initializeProps
End Sub

Private Sub initializeProps()
    Set pPropsByAliases = F.Dictionaries.Create(False)
    Set pPropsByName = F.Dictionaries.Create(False)
    Set pPropsById = F.Dictionaries.Create(False)
    '[Size]
    Call addProp(1, PROP_WIDTH, False, "width", "size.width").setIsSizeProperty(True)
    Call addProp(2, PROP_MIN_WIDTH, False, "min-width", "minWidth", "size.minWidth", "size.width.min").setIsSizeProperty(True)
    Call addProp(3, PROP_MAX_WIDTH, False, "max-width", "maxWidth", "size.maxWidth", "size.width.max").setIsSizeProperty(True)
    Call addProp(4, PROP_HEIGHT, False, "height", "size.height").setIsSizeProperty(True)
    Call addProp(5, PROP_MIN_HEIGHT, False, "min-height", "minHeight", "size.minHeight", "size.height.min").setIsSizeProperty(True)
    Call addProp(6, PROP_MAX_HEIGHT, False, "max-height", "maxHeight", "size.maxHeight", "size.height.max").setIsSizeProperty(True)
    '[Positions
    Call addProp(7, PROP_TOP, False, "top", "position.top").setIsPositionProperty(True)
    Call addProp(8, PROP_RIGHT, False, "right", "position.right").setIsPositionProperty(True)
    Call addProp(9, PROP_BOTTOM, False, "bottom", "position.bottom").setIsPositionProperty(True)
    Call addProp(10, PROP_LEFT, False, "left", "position.left").setIsPositionProperty(True)
    '[Margins & paddings]
    Call addProp(11, PROP_MARGIN, False, "margin").setIsPositionProperty(True)
    Call addProp(12, PROP_MARGIN_VERTICAL, False, "margin-vertical", "margin.vertical", "marginVertical", _
                        "vertical-margin", "vertical.margin", "vertical_margin", "verticalMargin").setIsPositionProperty(True)
    Call addProp(13, PROP_MARGIN_HORIZONTAL, False, "margin-horizontal", "margin.horizontal", "marginHorizontal", _
                        "horizontal-margin", "horizontal.margin", "horizontal_margin", "horizontalMargin").setIsPositionProperty(True)
    Call addProp(14, PROP_MARGIN_TOP, False, "margin-top", "marginTop", "top-margin", "topMargin", "margin.top").setIsPositionProperty(True)
    Call addProp(15, PROP_MARGIN_RIGHT, False, "margin-right", "marginRight", "right-margin", "rightMargin", "margin.right").setIsPositionProperty(True)
    Call addProp(16, PROP_MARGIN_BOTTOM, False, "margin-bottom", "marginBottom", "bottom-margin", "bottomMargin", "margin.bottom").setIsPositionProperty(True)
    Call addProp(17, PROP_MARGIN_LEFT, False, "margin-left", "marginLeft", "left-margin", "leftMargin", "margin.left").setIsPositionProperty(True)
    
    Call addProp(18, PROP_PADDING, False, "padding").setIsLayoutProperty(True)
    Call addProp(19, PROP_PADDING_VERTICAL, False, "padding-vertical", "padding.vertical", "paddingVertical", _
                        "vertical-padding", "vertical.padding", "vertical_padding", "verticalPadding").setIsLayoutProperty(True)
    Call addProp(20, PROP_PADDING_HORIZONTAL, False, "padding-horizontal", "padding.horizontal", "paddingHorizontal", _
                        "horizontal-padding", "horizontal.padding", "horizontal_padding", "horizontalPadding").setIsLayoutProperty(True)
    Call addProp(21, PROP_PADDING_TOP, False, "padding-top", "paddingTop", "top-padding", "topPadding", "padding.top").setIsLayoutProperty(True)
    Call addProp(22, PROP_PADDING_RIGHT, False, "padding-right", "paddingRight", "right-padding", "rightPadding", "padding.right").setIsLayoutProperty(True)
    Call addProp(23, PROP_PADDING_BOTTOM, False, "padding-bottom", "paddingBottom", "bottom-padding", "bottomPadding", "padding.bottom").setIsLayoutProperty(True)
    Call addProp(24, PROP_PADDING_LEFT, False, "padding-left", "paddingLeft", "left-padding", "leftPadding", "padding.left").setIsLayoutProperty(True)
    '[Background]
    Call addProp(25, PROP_BACKGROUND_STYLE, False, "backgroundStyle", "background-style", "background.style")
    Call addProp(26, PROP_BACKGROUND_COLOR, False, "backgroundColor", "background-color", "background.color", _
                        "backColor", "back.color", "back_color").setConvertionFunction("convertRgbToColor")
    '[Font]
    Call addProp(27, PROP_FONT_SIZE, True, "fontSize", "font-size", "font.size")
    Call addProp(28, PROP_FONT_COLOR, True, "fontColor", "font-color", "font.color").setConvertionFunction("convertRgbToColor")
    Call addProp(29, PROP_FONT_FAMILY, True, "fontName", "font-name", "font.name", "fontFamily", "font-family", "font.family")
    Call addProp(30, PROP_FONT_BOLD, True, "fontBold", "font-bold", "font.bold")
    Call addProp(31, PROP_FONT_ITALIC, True, "fontItalic", "font-italic", "font.italic")
    '[Border]
    Call addProp(32, PROP_BORDERS, False, "borders")
    Call addProp(33, PROP_BORDER_VISIBLE, False, "borderVisible", "border-visible", "border.visible")
    Call addProp(34, PROP_BORDER_COLOR, True, "borderColor", "border-color", "border.color").setConvertionFunction("convertRgbToColor")
    Call addProp(35, PROP_BORDER_WIDTH, False, "borderWidth", "border-width", "border.width", _
                        "borderSize", "border-size", "border.size", "borderWeight", "border-weight", "border.weight")
    Call addProp(36, PROP_BORDER_STYLE, False, "borderStyle", "border-style", "border.style")
    
    '[Text]
    Call addProp(37, PROP_HORIZONTAL_ALIGNMENT, True, _
                        "horizontalAlignment", "horizontal-alignment", "horizontal.alignment", "horizontal-align", "horizontalAlign", "horizontal.align", _
                        "hAlignment", "h-alignment", "h.alignment", "h-align", "hAlign", "h.align", "alignment.horizontal"). _
                        setConvertionFunction("convertAlignTextToEnum")
    Call addProp(38, PROP_VERTICAL_ALIGNMENT, True, _
                        "verticalAlignment", "vertical-alignment", "vertical.alignment", "vertical-align", "verticalAlign", "vertical.align", _
                        "vAlignment", "v-alignment", "v.alignment", "v-align", "vAlign", "v.align", "alignment.vertical"). _
                        setConvertionFunction("convertAlignTextToEnum")
    Call addProp(39, PROP_TEXT_UPPERCASE, False, "textUppercase", "text-uppercase", "text.uppercase")
    
    '[Picture]
    Call addProp(40, PROP_PICTURE, False, "picture")
    Call addProp(41, PROP_PICTURE_HEIGHT, False, "pictureHeight", "picture-height", "picture.height")
    Call addProp(42, PROP_PICTURE_WIDTH, False, "pictureWidth", "picture-width", "picture.width")
    Call addProp(43, PROP_PICTURE_MARGIN, False, "pictureMargin", "picture-margin", "picture.margin")
    
    '[Controls' specific]
    Call addProp(44, PROP_MATCH_REQUIRED, False, "matchRequired", "match-required", "match.required", "match_required")
    Call addProp(45, PROP_TRIPLE_STATE, False, "tripleState", "triple-state", "triple_state")
    
    '::List
    Call addProp(46, PROP_SCROLL_BARS, False, "scrollBars", "scroll-bars", "scroll_bars")
    Call addProp(47, PROP_LINE_HEIGHT, False, "lineHeight", "line-height", "line_height")
    'inheritable: text-align, text-transform, text-indent, line-height, border-color
    
    '[Worksheet]
    Call addProp(48, PROP_INDENT_LEVEL, False, "indent", "indent-level", "indent.level", "indent_level")
    Call addProp(49, PROP_WRAP_TEXT, False, "wrapText", "textWrap", "wrap-text", "wrap.text", _
                                            "wrap_text", "text-wrap", "text.wrap", "text_wrap", _
                                            "wrapWord", "wordWrap", "wrap-word", "wrap.word", _
                                            "wrap_word", "word-wrap", "word.wrap", "word_wrap")
    Call addProp(50, PROP_AUTO_FIT, False, "autoFit", "auto-fit", "auto_fit", "auto.fit")
    Call addProp(51, PROP_NUMBER_FORMAT, False, "numberFormat", "number-format", "number_format", "number.format")
    
End Sub

Private Function addProp(id As Long, propName As String, isInheritable As Boolean, ParamArray aliases() As Variant) As UIProp
    Dim varAlias As Variant
    '----------------------------------------------------------------------------------------------------------
    Set addProp = New UIProp
    With addProp
        Call .setId(id)
        Call .setName(propName)
        Call .setIsInheritable(isInheritable)
        '[Add to repositories]
        Call pPropsByName.Add(propName, addProp)
        Call pPropsById.Add(id, addProp)
        '[Add aliases]
        For Each varAlias In aliases
            Call addProp.addAlias(VBA.CStr(varAlias))
            Call F.Dictionaries.addItem(pPropsByAliases, varAlias, addProp, duplicateBehaviour_WarningInImmediateWindow)
        Next varAlias
    End With
End Function




'[API]
Public Function getPropById(id As Long) As UIProp
    Set getPropById = F.Dictionaries.getObject(pPropsById, id)
End Function

Public Function getPropByName(name As String) As UIProp
    Set getPropByName = F.Dictionaries.getObject(pPropsByName, name)
End Function

Public Function getPropByAlias(alias As String) As UIProp
    Set getPropByAlias = F.Dictionaries.getObject(pPropsByAliases, alias)
End Function

Public Function getPropsArray(Optional withDefaultValues As Boolean = True) As Variant
    If withDefaultValues Then
        getPropsArray = pDefaultPropsArray
    Else
        getPropsArray = pEmptyPropsArray
    End If
End Function






'[Quick Access]
'[Size]
Public Function minWidth() As UIProp: Set minWidth = F.Dictionaries.getObject(pPropsByAliases, PROP_MIN_WIDTH): End Function
Public Function width() As UIProp: Set width = F.Dictionaries.getObject(pPropsByAliases, PROP_WIDTH): End Function
Public Function maxWidth() As UIProp: Set maxWidth = F.Dictionaries.getObject(pPropsByAliases, PROP_MAX_WIDTH): End Function
Public Function minHeight() As UIProp: Set minHeight = F.Dictionaries.getObject(pPropsByAliases, PROP_MIN_HEIGHT): End Function
Public Function height() As UIProp: Set height = F.Dictionaries.getObject(pPropsByAliases, PROP_HEIGHT): End Function
Public Function maxHeight() As UIProp: Set maxHeight = F.Dictionaries.getObject(pPropsByAliases, PROP_MAX_HEIGHT): End Function
'[Position]
Public Function top() As UIProp: Set top = F.Dictionaries.getObject(pPropsByAliases, PROP_TOP): End Function
Public Function right() As UIProp: Set right = F.Dictionaries.getObject(pPropsByAliases, PROP_RIGHT): End Function
Public Function bottom() As UIProp: Set bottom = F.Dictionaries.getObject(pPropsByAliases, PROP_BOTTOM): End Function
Public Function left() As UIProp: Set left = F.Dictionaries.getObject(pPropsByAliases, PROP_LEFT): End Function
'[Margins & paddings]
Public Function Margin() As UIProp: Set Margin = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN): End Function
Public Function MarginVertical() As UIProp: Set MarginVertical = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_VERTICAL): End Function
Public Function MarginHorizontal() As UIProp: Set MarginHorizontal = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_HORIZONTAL): End Function
Public Function MarginTop() As UIProp: Set MarginTop = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_TOP): End Function
Public Function MarginRight() As UIProp: Set MarginRight = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_RIGHT): End Function
Public Function MarginBottom() As UIProp: Set MarginBottom = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_BOTTOM): End Function
Public Function MarginLeft() As UIProp: Set MarginLeft = F.Dictionaries.getObject(pPropsByAliases, PROP_MARGIN_LEFT): End Function
Public Function Padding() As UIProp: Set Padding = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING): End Function
Public Function PaddingVertical() As UIProp: Set PaddingVertical = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_VERTICAL): End Function
Public Function PaddingHorizontal() As UIProp: Set PaddingHorizontal = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_HORIZONTAL): End Function
Public Function PaddingTop() As UIProp: Set PaddingTop = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_TOP): End Function
Public Function PaddingRight() As UIProp: Set PaddingRight = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_RIGHT): End Function
Public Function PaddingBottom() As UIProp: Set PaddingBottom = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_BOTTOM): End Function
Public Function PaddingLeft() As UIProp: Set PaddingLeft = F.Dictionaries.getObject(pPropsByAliases, PROP_PADDING_LEFT): End Function
'[Background]
Public Function BackgroundStyle() As UIProp: Set BackgroundStyle = F.Dictionaries.getObject(pPropsByAliases, PROP_BACKGROUND_STYLE): End Function
Public Function BackgroundColor() As UIProp: Set BackgroundColor = F.Dictionaries.getObject(pPropsByAliases, PROP_BACKGROUND_COLOR): End Function
'[Font]
Public Function fontSize() As UIProp: Set fontSize = F.Dictionaries.getObject(pPropsByAliases, PROP_FONT_SIZE): End Function
Public Function FontColor() As UIProp: Set FontColor = F.Dictionaries.getObject(pPropsByAliases, PROP_FONT_COLOR): End Function
Public Function FontFamily() As UIProp: Set FontFamily = F.Dictionaries.getObject(pPropsByAliases, PROP_FONT_FAMILY): End Function
Public Function FontBold() As UIProp: Set FontBold = F.Dictionaries.getObject(pPropsByAliases, PROP_FONT_BOLD): End Function
Public Function FontItalic() As UIProp: Set FontItalic = F.Dictionaries.getObject(pPropsByAliases, PROP_FONT_ITALIC): End Function
'[Border]
Public Function Borders() As UIProp: Set Borders = F.Dictionaries.getObject(pPropsByAliases, PROP_BORDERS): End Function
Public Function BorderVisible() As UIProp: Set BorderVisible = F.Dictionaries.getObject(pPropsByAliases, PROP_BORDER_VISIBLE): End Function
Public Function BorderColor() As UIProp: Set BorderColor = F.Dictionaries.getObject(pPropsByAliases, PROP_BORDER_COLOR): End Function
Public Function borderWidth() As UIProp: Set borderWidth = F.Dictionaries.getObject(pPropsByAliases, PROP_BORDER_WIDTH): End Function
Public Function BorderStyle() As UIProp: Set BorderStyle = F.Dictionaries.getObject(pPropsByAliases, PROP_BORDER_STYLE): End Function
'[Text]
Public Function HorizontalAlignment() As UIProp: Set HorizontalAlignment = F.Dictionaries.getObject(pPropsByAliases, PROP_HORIZONTAL_ALIGNMENT): End Function
Public Function VerticalAlignment() As UIProp: Set VerticalAlignment = F.Dictionaries.getObject(pPropsByAliases, PROP_VERTICAL_ALIGNMENT): End Function
Public Function TextUppercase() As UIProp: Set TextUppercase = F.Dictionaries.getObject(pPropsByAliases, PROP_TEXT_UPPERCASE): End Function
'[Pictures]
Public Function Picture() As UIProp: Set Picture = F.Dictionaries.getObject(pPropsByAliases, PROP_PICTURE): End Function
Public Function PictureWidth() As UIProp: Set PictureWidth = F.Dictionaries.getObject(pPropsByAliases, PROP_PICTURE_WIDTH): End Function
Public Function PictureHeight() As UIProp: Set PictureHeight = F.Dictionaries.getObject(pPropsByAliases, PROP_PICTURE_HEIGHT): End Function
Public Function PictureMargin() As UIProp: Set PictureMargin = F.Dictionaries.getObject(pPropsByAliases, PROP_PICTURE_MARGIN): End Function
'[Combo boxes]
Public Function MatchRequired() As UIProp: Set MatchRequired = F.Dictionaries.getObject(pPropsByAliases, PROP_MATCH_REQUIRED): End Function
'[Checkbox]
Public Function TripleState() As UIProp: Set TripleState = F.Dictionaries.getObject(pPropsByAliases, PROP_TRIPLE_STATE): End Function
'[List]
Public Function LineHeight() As UIProp: Set LineHeight = F.Dictionaries.getObject(pPropsByAliases, PROP_LINE_HEIGHT): End Function
Public Function ScrollBars() As UIProp: Set ScrollBars = F.Dictionaries.getObject(pPropsByAliases, PROP_SCROLL_BARS): End Function
'[Worksheet]
Public Function IndentLevel() As UIProp: Set IndentLevel = F.Dictionaries.getObject(pPropsByAliases, PROP_INDENT_LEVEL): End Function
Public Function WrapText() As UIProp: Set WrapText = F.Dictionaries.getObject(pPropsByAliases, PROP_WRAP_TEXT): End Function
Public Function AutoFit() As UIProp: Set AutoFit = F.Dictionaries.getObject(pPropsByAliases, PROP_AUTO_FIT): End Function
Public Function NumberFormat() As UIProp: Set NumberFormat = F.Dictionaries.getObject(pPropsByAliases, PROP_NUMBER_FORMAT): End Function
