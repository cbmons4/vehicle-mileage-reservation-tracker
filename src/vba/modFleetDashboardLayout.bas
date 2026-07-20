Attribute VB_Name = "modFleetDashboardLayout"
Option Explicit

Private Const DASHBOARD_CHOICE_PREFIX As String = "dashChoice"

Public Sub FleetApplyDashboardPolish(ByVal ws As Worksheet)
    CleanDashboardCardCanvas ws
    AlignUpperHeaders ws
    AlignLowerHeaders ws
    AlignReservationCallout ws
    AlignCardOutlines ws
    ClearActionRowBackgrounds ws
    StylePrimaryActionButtons ws
    StyleCalendarButtons ws
    ws.Shapes("btnSAVECHECKOUTE11F11").TextFrame.Characters.Text = "SAVE RESERVATION"
    StyleNavigationShape ws.Shapes("btnMONTHLYMILEAGEA11D11")
    StyleNavigationShape ws.Shapes("btnWEEKLYCHECKOUTBOARDE12H15")
    CenterReservationsButton ws
    StyleReviewPanel ws
    AlignReviewSelectors ws
    AlignLegendShapes ws
    AlignHeaderIconBadges ws
    AlignDatePickerGlyphs ws
End Sub

Public Sub FleetRefreshDashboardReviewHeader(ByVal ws As Worksheet)
    StyleReviewPanel ws
    AlignReviewSelectors ws
    AlignHeaderIconBadges ws
End Sub

Private Sub CleanDashboardCardCanvas(ByVal ws As Worksheet)
    ClearHeaderBase ws.Range("A5:H5")
    ClearHeaderBase ws.Range("A12:H12")
    ClearCardExteriorEdges ws.Range("A6:D11")
    ClearCardExteriorEdges ws.Range("E6:H11")
    ClearCardExteriorEdges ws.Range("A13:D15")
    ClearCardExteriorEdges ws.Range("E13:H15")
End Sub

Private Sub ClearHeaderBase(ByVal targetRange As Range)
    targetRange.Interior.Color = RGB(255, 255, 255)
    targetRange.Borders.LineStyle = xlNone
End Sub

Private Sub ClearCardExteriorEdges(ByVal targetRange As Range)
    targetRange.Borders(xlEdgeLeft).LineStyle = xlNone
    targetRange.Borders(xlEdgeRight).LineStyle = xlNone
End Sub

Public Sub FleetOpenFiscalYearDropdown()
    ShowDashboardChoiceMenu "FiscalYearChoiceList", False, "FY", _
        "D17", "v5FiscalValueButton"
End Sub

Public Sub FleetOpenReportMonthDropdown()
    ShowDashboardChoiceMenu "MonthChoiceList", True, "MONTH", _
        "G17", "v5ReportMonthValueButton"
End Sub

Private Sub ShowDashboardChoiceMenu(ByVal choiceListName As String, ByVal formatAsMonth As Boolean, _
    ByVal menuType As String, ByVal fallbackCellAddress As String, ByVal anchorShapeName As String)

    Dim ws As Worksheet
    Dim choiceRange As Range
    Dim choiceCell As Range
    Dim anchorShape As Shape
    Dim panelShape As Shape
    Dim optionShape As Shape
    Dim choiceCaption As String
    Dim choiceTag As String
    Dim currentCaption As String
    Dim menuLeft As Double
    Dim menuTop As Double
    Dim menuWidth As Double
    Dim optionHeight As Double
    Dim optionIndex As Long
    Dim optionCount As Long
    Dim existingMenuType As String

    On Error GoTo UseCellDropdown
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    On Error Resume Next
    existingMenuType = ws.Shapes(DASHBOARD_CHOICE_PREFIX & "Panel").AlternativeText
    On Error GoTo UseCellDropdown
    If StrComp(existingMenuType, menuType, vbTextCompare) = 0 Then
        FleetDismissDashboardChoicePopup
        Exit Sub
    End If

    FleetDismissDashboardDateCalendar ws.Range("A1")
    FleetDismissDashboardChoicePopup
    ThisWorkbook.Activate
    ws.Activate
    Set anchorShape = ws.Shapes(anchorShapeName)
    Set choiceRange = ThisWorkbook.names(choiceListName).RefersToRange
    currentCaption = Trim$(CStr(ws.Range(fallbackCellAddress).MergeArea.Cells(1, 1).Text))
    menuLeft = anchorShape.Left
    menuTop = anchorShape.Top + anchorShape.Height + 1
    menuWidth = anchorShape.Width
    optionHeight = 19

    For Each choiceCell In choiceRange.Cells
        If Len(Trim$(CStr(choiceCell.Value2))) > 0 Then
            If formatAsMonth And IsDate(choiceCell.value) Then
                choiceCaption = Format$(CDate(choiceCell.value), "mmmm yyyy")
                choiceTag = Format$(CDate(choiceCell.value), "yyyy-mm-dd")
            Else
                choiceCaption = Trim$(CStr(choiceCell.value))
                choiceTag = choiceCaption
            End If

            optionCount = optionCount + 1
            Set optionShape = ws.Shapes.AddShape(msoShapeRectangle, menuLeft + 1, _
                menuTop + 1 + ((optionCount - 1) * optionHeight), menuWidth - 2, optionHeight)
            With optionShape
                .Name = DASHBOARD_CHOICE_PREFIX & "Option" & Format$(optionCount, "00")
                .AlternativeText = menuType & "|" & choiceTag
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                If StrComp(choiceCaption, currentCaption, vbTextCompare) = 0 Then _
                    .Fill.ForeColor.RGB = RGB(255, 242, 204)
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = RGB(220, 224, 228)
                .Line.Weight = 0.5
                .TextFrame.Characters.Text = choiceCaption
                .TextFrame.Characters.Font.Name = "Aptos"
                .TextFrame.Characters.Font.Size = 9.5
                .TextFrame.Characters.Font.Color = RGB(25, 30, 35)
                .TextFrame.Characters.Font.Bold = False
                .TextFrame.HorizontalAlignment = xlLeft
                .TextFrame.VerticalAlignment = xlCenter
                .TextFrame.MarginLeft = 8
                .TextFrame.MarginRight = 4
                .TextFrame.MarginTop = 0
                .TextFrame.MarginBottom = 0
                .Placement = xlFreeFloating
                .OnAction = "FleetChooseDashboardPopupOption"
                .ZOrder msoBringToFront
            End With
        End If
    Next choiceCell

    If optionCount = 0 Then GoTo UseCellDropdown
    Set panelShape = ws.Shapes.AddShape(msoShapeRectangle, menuLeft, menuTop, _
        menuWidth, (optionCount * optionHeight) + 2)
    With panelShape
        .Name = DASHBOARD_CHOICE_PREFIX & "Panel"
        .AlternativeText = menuType
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(170, 176, 182)
        .Line.Weight = 0.75
        .Placement = xlFreeFloating
        .OnAction = "FleetDismissDashboardChoicePopup"
        .ZOrder msoBringToFront
    End With
    For optionIndex = 1 To optionCount
        ws.Shapes(DASHBOARD_CHOICE_PREFIX & "Option" & Format$(optionIndex, "00")).ZOrder msoBringToFront
    Next optionIndex
    Exit Sub

UseCellDropdown:
    FleetDismissDashboardChoicePopup
    On Error Resume Next
    ws.Activate
    ws.Range(fallbackCellAddress).Select
    On Error GoTo 0
End Sub

Public Sub FleetChooseDashboardPopupOption()
    On Error Resume Next
    FleetApplyDashboardPopupShapeChoice CStr(Application.Caller)
    On Error GoTo 0
End Sub

Public Sub FleetApplyDashboardPopupShapeChoice(ByVal shapeName As String)
    Dim ws As Worksheet
    Dim choiceTag As String
    Dim separatorPosition As Long
    Dim menuType As String
    Dim choiceValue As String

    On Error GoTo CleanExit
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    choiceTag = CStr(ws.Shapes(shapeName).AlternativeText)
    separatorPosition = InStr(1, choiceTag, "|", vbBinaryCompare)
    If separatorPosition = 0 Then GoTo CleanExit
    menuType = Left$(choiceTag, separatorPosition - 1)
    choiceValue = Mid$(choiceTag, separatorPosition + 1)

    FleetDismissDashboardChoicePopup
    If StrComp(menuType, "FY", vbTextCompare) = 0 Then
        FleetApplyDashboardFiscalYearChoice choiceValue
    ElseIf StrComp(menuType, "MONTH", vbTextCompare) = 0 Then
        FleetApplyDashboardReportMonthChoice choiceValue
    End If

CleanExit:
End Sub

Public Sub FleetApplyDashboardFiscalYearChoice(ByVal fiscalYearLabel As String)
    ApplyDashboardSelectorChoice "D17", Trim$(fiscalYearLabel), False
End Sub

Public Sub FleetApplyDashboardReportMonthChoice(ByVal monthIsoDate As String)
    ApplyDashboardSelectorChoice "G17", monthIsoDate, True
End Sub

Private Sub ApplyDashboardSelectorChoice(ByVal cellAddress As String, ByVal choiceValue As String, _
    ByVal valueIsIsoDate As Boolean)

    Dim ws As Worksheet
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim selectedValue As Variant

    On Error GoTo CleanExit
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    If Len(choiceValue) = 0 Then Exit Sub
    If valueIsIsoDate Then
        If Len(choiceValue) <> 10 Then Exit Sub
        selectedValue = DateSerial(CLng(Left$(choiceValue, 4)), _
            CLng(Mid$(choiceValue, 6, 2)), CLng(Right$(choiceValue, 2)))
    Else
        selectedValue = choiceValue
    End If

    FleetDismissDashboardChoicePopup
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range(cellAddress).MergeArea.Cells(1, 1).value = selectedValue
    FleetRefreshDashboard False

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
End Sub

Public Sub FleetDismissDashboardChoicePopup()
    Dim ws As Worksheet
    Dim shapeIndex As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        If InStr(1, ws.Shapes(shapeIndex).Name, DASHBOARD_CHOICE_PREFIX, vbTextCompare) = 1 Then _
            ws.Shapes(shapeIndex).Delete
    Next shapeIndex
    On Error GoTo 0
End Sub

Private Sub ClearActionRowBackgrounds(ByVal ws As Worksheet)
    ClearActionBackgroundRange ws.Range("A10:D11")
    ClearActionBackgroundRange ws.Range("E11:H11")
    ClearActionBackgroundRange ws.Range("A15:D15")
    ClearActionBackgroundRange ws.Range("E15:H15")
End Sub

Private Sub ClearActionBackgroundRange(ByVal targetRange As Range)
    targetRange.Interior.Color = RGB(255, 255, 255)
    targetRange.Borders.LineStyle = xlNone
End Sub

Private Sub StylePrimaryActionButtons(ByVal ws As Worksheet)
    PositionButtonPair ws.Range("A10:D10"), ws.Shapes("btnSAVEMILEAGEA10B10"), ws.Shapes("btnCLEARC10D10")
    PositionButtonPair ws.Range("E11:H11"), ws.Shapes("btnSAVECHECKOUTE11F11"), ws.Shapes("btnCLEARG11H11")
    PositionButtonPair ws.Range("A15:D15"), ws.Shapes("btnMARKINSHOPA15B15"), ws.Shapes("btnMARKAVAILABLEC15D15")

    StylePrimaryButton ws.Shapes("btnSAVEMILEAGEA10B10"), RGB(0, 122, 45), RGB(0, 92, 34), RGB(255, 255, 255)
    StylePrimaryButton ws.Shapes("btnSAVECHECKOUTE11F11"), RGB(0, 122, 45), RGB(0, 92, 34), RGB(255, 255, 255)
    StylePrimaryButton ws.Shapes("btnMARKINSHOPA15B15"), RGB(167, 25, 31), RGB(126, 18, 23), RGB(255, 255, 255)
    StylePrimaryButton ws.Shapes("btnMARKAVAILABLEC15D15"), RGB(0, 122, 45), RGB(0, 92, 34), RGB(255, 255, 255)
    StylePrimaryButton ws.Shapes("btnCLEARC10D10"), RGB(220, 224, 228), RGB(186, 193, 199), RGB(25, 30, 35)
    StylePrimaryButton ws.Shapes("btnCLEARG11H11"), RGB(220, 224, 228), RGB(186, 193, 199), RGB(25, 30, 35)
End Sub

Private Sub PositionButtonPair(ByVal targetRange As Range, ByVal leftButton As Shape, ByVal rightButton As Shape)
    Dim sideInset As Double
    Dim buttonGap As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim buttonTop As Double

    sideInset = 10
    buttonGap = 8
    buttonWidth = (targetRange.Width - (2 * sideInset) - buttonGap) / 2
    buttonHeight = targetRange.Height - 8
    If buttonHeight > 28 Then buttonHeight = 28
    buttonTop = targetRange.Top + ((targetRange.Height - buttonHeight) / 2)

    leftButton.Left = targetRange.Left + sideInset
    leftButton.Top = buttonTop
    leftButton.Width = buttonWidth
    leftButton.Height = buttonHeight
    leftButton.Placement = xlMove

    rightButton.Left = leftButton.Left + buttonWidth + buttonGap
    rightButton.Top = buttonTop
    rightButton.Width = buttonWidth
    rightButton.Height = buttonHeight
    rightButton.Placement = xlMove
End Sub

Private Sub StylePrimaryButton(ByVal buttonShape As Shape, ByVal fillColor As Long, _
    ByVal lineColor As Long, ByVal fontColor As Long)

    buttonShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius buttonShape, 0.08
    buttonShape.Fill.ForeColor.RGB = fillColor
    buttonShape.Fill.Transparency = 0
    buttonShape.Line.Visible = msoTrue
    buttonShape.Line.ForeColor.RGB = lineColor
    buttonShape.Line.Weight = 1
    buttonShape.TextFrame.Characters.Font.Name = "Aptos"
    buttonShape.TextFrame.Characters.Font.Size = 9.5
    buttonShape.TextFrame.Characters.Font.Bold = True
    buttonShape.TextFrame.Characters.Font.Color = fontColor
    buttonShape.TextFrame.HorizontalAlignment = xlCenter
    buttonShape.TextFrame.VerticalAlignment = xlCenter
    buttonShape.Shadow.Visible = msoFalse
End Sub

Private Sub StyleCalendarButtons(ByVal ws As Worksheet)
    PositionCalendarButton ws.Shapes("btnCALD6D6"), ws.Range("D6")
    PositionCalendarButton ws.Shapes("btnCALH6H6"), ws.Range("H6")
End Sub

Private Sub PositionCalendarButton(ByVal buttonShape As Shape, ByVal targetCell As Range)
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    buttonWidth = 56
    buttonHeight = 24
    buttonShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius buttonShape, 0.08
    buttonShape.Left = targetCell.Left + ((targetCell.Width - buttonWidth) / 2)
    buttonShape.Top = targetCell.Top + ((targetCell.Height - buttonHeight) / 2)
    buttonShape.Width = buttonWidth
    buttonShape.Height = buttonHeight
    buttonShape.Fill.ForeColor.RGB = RGB(0, 43, 85)
    buttonShape.Fill.Transparency = 0
    buttonShape.Line.Visible = msoTrue
    buttonShape.Line.ForeColor.RGB = RGB(0, 34, 68)
    buttonShape.Line.Weight = 1
    buttonShape.TextFrame.Characters.Font.Name = "Aptos"
    buttonShape.TextFrame.Characters.Font.Size = 9.5
    buttonShape.TextFrame.Characters.Font.Bold = True
    buttonShape.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    buttonShape.TextFrame.HorizontalAlignment = xlCenter
    buttonShape.TextFrame.VerticalAlignment = xlCenter
    buttonShape.Shadow.Visible = msoFalse
    buttonShape.Placement = xlMove
End Sub

Private Sub SetSubtleCornerRadius(ByVal targetShape As Shape, ByVal radiusValue As Double)
    On Error Resume Next
    targetShape.Adjustments.item(1) = radiusValue
    On Error GoTo 0
End Sub

Private Sub AlignHeaderIconBadges(ByVal ws As Worksheet)
    PositionExactHeaderIcon ws, "v5LogHeaderCircle", "v3LogHeaderIcon", "v2LogHeader", 24, 5
    PositionExactHeaderIcon ws, "v5ReserveHeaderCircle", "v3ReserveHeaderIcon", "v2CheckoutHeader", 24, 10
    PositionExactHeaderIcon ws, "v5ReservationsHeaderCircle", "v3ReservationsHeaderIcon", "v2ReservationsHeader", 24, 10
    PositionExactHeaderIcon ws, "v5ShopHeaderCircle", "v3ShopHeaderIcon", "v2ShopHeader", 24, 5
    PositionExactHeaderIcon ws, "v5ReviewHeaderCircle", "v3ReviewHeaderIcon", "v2ReviewTitleBand", 24, 6
End Sub

Private Sub PositionExactHeaderIcon(ByVal ws As Worksheet, ByVal circleName As String, _
    ByVal iconName As String, ByVal headerName As String, ByVal iconSize As Double, _
    ByVal leftInset As Double)

    Dim circleShape As Shape
    Dim iconShape As Shape
    Dim headerShape As Shape

    On Error Resume Next
    Set circleShape = ws.Shapes(circleName)
    Set iconShape = ws.Shapes(iconName)
    Set headerShape = ws.Shapes(headerName)
    On Error GoTo 0
    If circleShape Is Nothing Or iconShape Is Nothing Or headerShape Is Nothing Then Exit Sub

    circleShape.Visible = msoFalse
    iconShape.LockAspectRatio = msoFalse
    iconShape.Left = headerShape.Left + leftInset
    iconShape.Top = headerShape.Top + ((headerShape.Height - iconSize) / 2)
    iconShape.Width = iconSize
    iconShape.Height = iconSize
    iconShape.Placement = xlFreeFloating
    iconShape.ZOrder msoBringToFront
End Sub

Private Sub PositionHeaderIconBadge(ByVal ws As Worksheet, ByVal circleName As String, _
    ByVal iconName As String, ByVal headerName As String, ByVal circleSize As Double, _
    ByVal iconSize As Double, ByVal leftInset As Double)

    Dim circleShape As Shape
    Dim iconShape As Shape
    Dim headerShape As Shape

    On Error Resume Next
    Set circleShape = ws.Shapes(circleName)
    Set iconShape = ws.Shapes(iconName)
    Set headerShape = ws.Shapes(headerName)
    On Error GoTo 0
    If circleShape Is Nothing Or iconShape Is Nothing Or headerShape Is Nothing Then Exit Sub

    circleShape.AutoShapeType = msoShapeOval
    circleShape.Left = headerShape.Left + leftInset
    circleShape.Top = headerShape.Top + ((headerShape.Height - circleSize) / 2)
    circleShape.Width = circleSize
    circleShape.Height = circleSize
    circleShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    circleShape.Fill.Transparency = 0
    circleShape.Line.Visible = msoFalse
    circleShape.Placement = xlFreeFloating
    circleShape.ZOrder msoBringToFront

    iconShape.LockAspectRatio = msoFalse
    iconShape.Left = circleShape.Left + ((circleSize - iconSize) / 2)
    iconShape.Top = circleShape.Top + ((circleSize - iconSize) / 2)
    iconShape.Width = iconSize
    iconShape.Height = iconSize
    iconShape.Placement = xlFreeFloating
    iconShape.ZOrder msoBringToFront
End Sub

Private Sub AlignDatePickerGlyphs(ByVal ws As Worksheet)
    PositionDatePickerGlyph ws, "v3MileageDateIcon", "v5MileageDateSepLeft", "v5MileageDateSepRight", "C6", "FleetOpenMileageDateCalendar"
    PositionDatePickerGlyph ws, "v3CheckoutDateIcon", "v5CheckoutDateSepLeft", "v5CheckoutDateSepRight", "G6", "FleetOpenCheckoutDateCalendar"
End Sub

Private Sub PositionDatePickerGlyph(ByVal ws As Worksheet, ByVal iconName As String, _
    ByVal leftSeparatorName As String, ByVal rightSeparatorName As String, _
    ByVal cellAddress As String, ByVal macroName As String)

    Dim targetCell As Range
    Dim iconShape As Shape
    Dim leftSeparator As Shape
    Dim rightSeparator As Shape
    Dim centerX As Double
    Dim separatorTop As Double

    On Error Resume Next
    Set targetCell = ws.Range(cellAddress)
    Set iconShape = ws.Shapes(iconName)
    Set leftSeparator = ws.Shapes(leftSeparatorName)
    Set rightSeparator = ws.Shapes(rightSeparatorName)
    On Error GoTo 0
    If iconShape Is Nothing Or leftSeparator Is Nothing Or rightSeparator Is Nothing Then Exit Sub

    centerX = targetCell.Left + targetCell.Width - 24
    iconShape.LockAspectRatio = msoFalse
    iconShape.Left = centerX - 7.5
    iconShape.Top = targetCell.Top + ((targetCell.Height - 15) / 2)
    iconShape.Width = 15
    iconShape.Height = 15
    iconShape.Placement = xlFreeFloating
    iconShape.OnAction = macroName
    iconShape.ZOrder msoBringToFront

    separatorTop = targetCell.Top + ((targetCell.Height - 18) / 2)
    StyleDateSeparator leftSeparator, centerX - 12, separatorTop, macroName
    StyleDateSeparator rightSeparator, centerX + 12, separatorTop, macroName
End Sub

Private Sub StyleDateSeparator(ByVal separatorShape As Shape, ByVal leftPosition As Double, _
    ByVal topPosition As Double, ByVal macroName As String)

    separatorShape.Left = leftPosition
    separatorShape.Top = topPosition
    separatorShape.Width = 0
    separatorShape.Height = 18
    separatorShape.Line.Visible = msoTrue
    separatorShape.Line.ForeColor.RGB = RGB(215, 220, 225)
    separatorShape.Line.Weight = 0.75
    separatorShape.Placement = xlFreeFloating
    separatorShape.OnAction = macroName
    separatorShape.ZOrder msoBringToFront
End Sub

Private Sub AlignReviewSelectors(ByVal ws As Worksheet)
    Dim fiscalRange As Range
    Dim monthRange As Range
    Dim fiscalLabel As Shape
    Dim fiscalValue As Shape
    Dim monthLabel As Shape
    Dim monthValue As Shape
    Dim labelWidth As Double
    Dim controlHeight As Double
    Dim controlTop As Double

    On Error Resume Next
    Set fiscalLabel = ws.Shapes("v5FiscalLabel")
    Set fiscalValue = ws.Shapes("v5FiscalValueButton")
    Set monthLabel = ws.Shapes("v5ReportMonthLabel")
    Set monthValue = ws.Shapes("v5ReportMonthValueButton")
    On Error GoTo 0
    If fiscalLabel Is Nothing Or fiscalValue Is Nothing Or monthLabel Is Nothing Or monthValue Is Nothing Then Exit Sub

    Set fiscalRange = ws.Range("D17").MergeArea
    Set monthRange = ws.Range("G17").MergeArea
    labelWidth = 66
    controlHeight = 24
    controlTop = fiscalRange.Top + ((fiscalRange.Height - controlHeight) / 2)

    PositionSelectorShape fiscalLabel, fiscalRange.Left + 2 - labelWidth, controlTop, labelWidth, controlHeight
    PositionSelectorShape fiscalValue, fiscalRange.Left + 2, controlTop, fiscalRange.Width - 2, controlHeight
    PositionSelectorShape monthLabel, monthRange.Left + 2 - labelWidth, controlTop, labelWidth, controlHeight
    PositionSelectorShape monthValue, monthRange.Left + 2, controlTop, monthRange.Width - 4, controlHeight

    StyleSelectorLabel fiscalLabel, "Fiscal Year"
    StyleSelectorValue fiscalValue, Trim$(CStr(ws.Range("D17").value)), "FleetOpenFiscalYearDropdown"
    StyleSelectorLabel monthLabel, "Report Month"
    StyleSelectorValue monthValue, Format$(ws.Range("G17").value, "mmmm yyyy"), "FleetOpenReportMonthDropdown"

    PositionSelectorCue ws, "v4FiscalDropdownCue", fiscalValue, "FleetOpenFiscalYearDropdown"
    PositionSelectorCue ws, "v4MonthDropdownCue", monthValue, "FleetOpenReportMonthDropdown"
End Sub

Private Sub AlignCardOutlines(ByVal ws As Worksheet)
    PositionCardOutline ws, "v6LogCardOutline", "v2LogHeader", ws.Range("A11:D11")
    PositionCardOutline ws, "v6ReserveCardOutline", "v2CheckoutHeader", ws.Range("E11:H11")
    PositionCardOutline ws, "v6ShopCardOutline", "v2ShopHeader", ws.Range("A15:D15")
    PositionCardOutline ws, "v6ReservationsCardOutline", "v2ReservationsHeader", ws.Range("E15:H15")
End Sub

Private Sub PositionCardOutline(ByVal ws As Worksheet, ByVal outlineName As String, _
    ByVal headerName As String, ByVal bottomRange As Range)

    Dim outlineShape As Shape
    Dim headerShape As Shape

    On Error Resume Next
    Set outlineShape = ws.Shapes(outlineName)
    Set headerShape = ws.Shapes(headerName)
    On Error GoTo 0
    If outlineShape Is Nothing Or headerShape Is Nothing Then Exit Sub

    outlineShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius outlineShape, 0.06
    outlineShape.Left = bottomRange.Left
    outlineShape.Top = headerShape.Top
    outlineShape.Width = bottomRange.Width
    outlineShape.Height = (bottomRange.Top + bottomRange.Height) - headerShape.Top
    outlineShape.Fill.Visible = msoFalse
    outlineShape.Line.Visible = msoTrue
    outlineShape.Line.ForeColor.RGB = RGB(215, 220, 225)
    outlineShape.Line.Weight = 1
    outlineShape.Placement = xlFreeFloating
    outlineShape.ZOrder msoSendToBack
End Sub

Private Sub AlignUpperHeaders(ByVal ws As Worksheet)
    PositionUpperHeader ws.Shapes("v2LogHeader"), ws.Range("A5:D5")
    PositionUpperHeader ws.Shapes("v2CheckoutHeader"), ws.Range("E5:H5")
End Sub

Private Sub PositionUpperHeader(ByVal headerShape As Shape, ByVal targetRange As Range)
    headerShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius headerShape, 0.06
    headerShape.Left = targetRange.Left + 1
    headerShape.Top = targetRange.Top + 1
    headerShape.Width = targetRange.Width - 2
    headerShape.Height = targetRange.Height - 2
    headerShape.Fill.ForeColor.RGB = RGB(0, 43, 85)
    headerShape.Line.Visible = msoFalse
    headerShape.Placement = xlFreeFloating
End Sub

Private Sub PositionSelectorShape(ByVal selectorShape As Shape, ByVal leftPosition As Double, _
    ByVal topPosition As Double, ByVal selectorWidth As Double, ByVal selectorHeight As Double)

    selectorShape.Left = leftPosition
    selectorShape.Top = topPosition
    selectorShape.Width = selectorWidth
    selectorShape.Height = selectorHeight
    selectorShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius selectorShape, 0.08
    selectorShape.Placement = xlFreeFloating
    selectorShape.ZOrder msoBringToFront
End Sub

Private Sub StyleSelectorLabel(ByVal selectorShape As Shape, ByVal caption As String)
    selectorShape.Fill.ForeColor.RGB = RGB(245, 246, 248)
    selectorShape.Fill.Transparency = 0
    selectorShape.Line.Visible = msoTrue
    selectorShape.Line.ForeColor.RGB = RGB(215, 220, 225)
    selectorShape.Line.Weight = 0.75
    selectorShape.TextFrame.Characters.Text = caption
    selectorShape.TextFrame.Characters.Font.Name = "Aptos"
    selectorShape.TextFrame.Characters.Font.Size = 9
    selectorShape.TextFrame.Characters.Font.Bold = True
    selectorShape.TextFrame.Characters.Font.Color = RGB(25, 30, 35)
    selectorShape.TextFrame.HorizontalAlignment = xlCenter
    selectorShape.TextFrame.VerticalAlignment = xlCenter
    selectorShape.TextFrame.MarginLeft = 2
    selectorShape.TextFrame.MarginRight = 2
    On Error Resume Next
    selectorShape.TextFrame2.WordWrap = msoFalse
    On Error GoTo 0
    selectorShape.OnAction = vbNullString
End Sub

Private Sub StyleSelectorValue(ByVal selectorShape As Shape, ByVal caption As String, ByVal macroName As String)
    selectorShape.Fill.ForeColor.RGB = RGB(255, 242, 204)
    selectorShape.Fill.Transparency = 0
    selectorShape.Line.Visible = msoTrue
    selectorShape.Line.ForeColor.RGB = RGB(215, 220, 225)
    selectorShape.Line.Weight = 0.75
    selectorShape.TextFrame.Characters.Text = caption
    selectorShape.TextFrame.Characters.Font.Name = "Aptos"
    selectorShape.TextFrame.Characters.Font.Size = 10
    selectorShape.TextFrame.Characters.Font.Bold = True
    selectorShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    selectorShape.TextFrame.HorizontalAlignment = xlCenter
    selectorShape.TextFrame.VerticalAlignment = xlCenter
    selectorShape.TextFrame.MarginLeft = 2
    selectorShape.TextFrame.MarginRight = 2
    On Error Resume Next
    selectorShape.TextFrame2.WordWrap = msoFalse
    On Error GoTo 0
    selectorShape.OnAction = macroName
End Sub

Private Sub PositionSelectorCue(ByVal ws As Worksheet, ByVal cueName As String, _
    ByVal valueShape As Shape, ByVal macroName As String)

    Dim cueShape As Shape

    On Error Resume Next
    Set cueShape = ws.Shapes(cueName)
    On Error GoTo 0
    If cueShape Is Nothing Then Exit Sub

    cueShape.Left = valueShape.Left + valueShape.Width - 16
    cueShape.Top = valueShape.Top + ((valueShape.Height - 10) / 2)
    cueShape.Width = 8
    cueShape.Height = 10
    cueShape.TextFrame.Characters.Font.Size = 7
    cueShape.Placement = xlFreeFloating
    cueShape.OnAction = macroName
    cueShape.ZOrder msoBringToFront
End Sub

Private Sub CenterReservationsButton(ByVal ws As Worksheet)
    Dim buttonShape As Shape
    Dim headerShape As Shape

    Set buttonShape = ws.Shapes("btnWEEKLYCHECKOUTBOARDE12H15")
    Set headerShape = ws.Shapes("v2ReservationsHeader")
    buttonShape.Left = (headerShape.Left + (headerShape.Width / 2)) - (buttonShape.Width / 2)
    buttonShape.Top = ws.Range("E15:H15").Top + 1
    buttonShape.Height = 26
End Sub

Private Sub AlignLowerHeaders(ByVal ws As Worksheet)
    Dim headerShape As Shape

    Set headerShape = ws.Shapes("v2ShopHeader")
    PositionLowerHeader headerShape, ws.Range("A12:D12")
    Set headerShape = ws.Shapes("v2ReservationsHeader")
    PositionLowerHeader headerShape, ws.Range("E12:H12")
End Sub

Private Sub PositionLowerHeader(ByVal headerShape As Shape, ByVal targetRange As Range)
    headerShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius headerShape, 0.06
    headerShape.Left = targetRange.Left + 1
    headerShape.Top = targetRange.Top + 6
    headerShape.Width = targetRange.Width - 2
    headerShape.Height = targetRange.Height - 7
    headerShape.Fill.ForeColor.RGB = RGB(0, 43, 85)
    headerShape.Line.Visible = msoFalse
    headerShape.Placement = xlFreeFloating
End Sub

Private Sub AlignReservationCallout(ByVal ws As Worksheet)
    Dim bodyRange As Range
    Dim iconShape As Shape
    Dim textShape As Shape
    Dim groupLeft As Double

    Set bodyRange = ws.Range("E13:H14")
    groupLeft = bodyRange.Left + ((bodyRange.Width - 228) / 2)

    Set iconShape = ws.Shapes("v3ReservationBodyIcon")
    iconShape.LockAspectRatio = msoFalse
    iconShape.Left = groupLeft
    iconShape.Top = bodyRange.Top + 8
    iconShape.Width = 46
    iconShape.Height = 46
    iconShape.Placement = xlMove
    iconShape.ZOrder msoBringToFront

    Set textShape = ws.Shapes("v2ReservationBodyText")
    textShape.Left = groupLeft + 58
    textShape.Top = bodyRange.Top + 8
    textShape.Width = 170
    textShape.Height = 46
    textShape.TextFrame.Characters.Font.Name = "Aptos"
    textShape.TextFrame.Characters.Font.Size = 11
    textShape.TextFrame.HorizontalAlignment = xlLeft
    textShape.TextFrame.VerticalAlignment = xlCenter
    textShape.Placement = xlMove
    textShape.ZOrder msoBringToFront
End Sub

Private Sub AlignLegendShapes(ByVal ws As Worksheet)
    Dim legendRange As Range
    Dim backShape As Shape
    Dim titleShape As Shape
    Dim swatchShape As Shape
    Dim labelShape As Shape
    Dim baseLeft As Double
    Dim topPosition As Double
    Dim positions As Variant
    Dim swatchColors As Variant
    Dim shapeIndex As Long

    Set legendRange = ws.Range("A16:H16")
    Set backShape = ws.Shapes("v2LegendBack")
    backShape.Width = 430
    backShape.Height = 20
    backShape.Left = legendRange.Left + ((legendRange.Width - backShape.Width) / 2)
    backShape.Top = legendRange.Top + 7
    backShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    backShape.Fill.Transparency = 0
    backShape.Line.Visible = msoTrue
    backShape.Line.ForeColor.RGB = RGB(205, 211, 217)
    backShape.Line.Weight = 0.75
    backShape.Placement = xlMove

    baseLeft = backShape.Left + 14
    topPosition = backShape.Top + 2
    Set titleShape = ws.Shapes("v2LegendTitle")
    titleShape.Left = baseLeft
    titleShape.Top = topPosition
    titleShape.Width = 88
    titleShape.Height = 16
    titleShape.TextFrame.Characters.Font.Name = "Aptos"
    titleShape.TextFrame.Characters.Font.Size = 9
    titleShape.TextFrame.Characters.Font.Bold = True
    titleShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    titleShape.Placement = xlFreeFloating

    positions = Array(100, 195, 290)
    swatchColors = Array(RGB(210, 228, 248), RGB(213, 235, 207), RGB(250, 205, 203))
    For shapeIndex = 0 To 2
        Set swatchShape = ws.Shapes("v2LegendSwatch" & CStr(shapeIndex + 1))
        swatchShape.Left = baseLeft + positions(shapeIndex)
        swatchShape.Top = topPosition + 2
        swatchShape.Width = 20
        swatchShape.Height = 12
        swatchShape.Fill.ForeColor.RGB = swatchColors(shapeIndex)
        swatchShape.Fill.Transparency = 0
        swatchShape.Line.Visible = msoTrue
        swatchShape.Line.ForeColor.RGB = swatchColors(shapeIndex)
        swatchShape.Placement = xlMove

        Set labelShape = ws.Shapes("v2LegendLabel" & CStr(shapeIndex + 1))
        labelShape.Left = baseLeft + positions(shapeIndex) + 26
        labelShape.Top = topPosition
        labelShape.Width = 62
        labelShape.Height = 16
        labelShape.TextFrame.Characters.Font.Name = "Aptos"
        labelShape.TextFrame.Characters.Font.Size = 9
        labelShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        labelShape.Placement = xlMove
    Next shapeIndex
End Sub

Private Sub StyleReviewPanel(ByVal ws As Worksheet)
    Dim titleShape As Shape

    Set titleShape = ws.Shapes("v2ReviewTitleBand")
    titleShape.Left = ws.Range("A17").Left
    titleShape.Top = ws.Range("A17").Top
    titleShape.Width = (ws.Range("D17").Left + 2 - 66) - titleShape.Left
    titleShape.Height = ws.Range("A17:H17").Height
    titleShape.AutoShapeType = msoShapeRectangle
    titleShape.Fill.Visible = msoFalse
    titleShape.Line.Visible = msoFalse
    titleShape.TextFrame.Characters.Text = "MONTHLY MILEAGE REVIEW"
    titleShape.TextFrame.Characters.Font.Name = "Aptos Display"
    titleShape.TextFrame.Characters.Font.Size = 10.5
    titleShape.TextFrame.Characters.Font.Bold = True
    titleShape.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    titleShape.TextFrame.HorizontalAlignment = xlLeft
    titleShape.TextFrame.VerticalAlignment = xlCenter
    titleShape.TextFrame.MarginLeft = 36
    On Error Resume Next
    titleShape.TextFrame2.WordWrap = msoFalse
    titleShape.TextFrame2.AutoSize = msoAutoSizeNone
    On Error GoTo 0
    titleShape.Placement = xlFreeFloating

    With ws.Range("A17:H17")
        .Interior.Color = RGB(0, 43, 85)
        .Font.Color = RGB(0, 43, 85)
        .Borders.LineStyle = xlNone
    End With
    With ws.Range("A17:H17").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RGB(215, 220, 225)
        .Weight = xlThin
    End With
    With ws.Range("A17:H17").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(215, 220, 225)
        .Weight = xlThin
    End With
    With ws.Range("A17:H17").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(215, 220, 225)
        .Weight = xlThin
    End With

    ws.Rows(19).RowHeight = 38
    With ws.Range("C20:C23")
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(24).RowHeight = 30
    With ws.Range("A19:H19")
        .Interior.Color = RGB(0, 72, 112)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A20:H24")
        .Font.Name = "Aptos"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Range("C20:C23").Font.Size = 9
    ws.Range("D20:H24").HorizontalAlignment = xlCenter
    ws.Range("B20:B23").HorizontalAlignment = xlRight
    StyleDashboardReviewGrid ws
    AutoFitDashboardReviewRows ws
    With ws.Range("A20:A24")
        .WrapText = False
        .ShrinkToFit = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub StyleDashboardReviewGrid(ByVal ws As Worksheet)
    With ws.Range("A19:H19").Borders
        .LineStyle = xlContinuous
        .Color = RGB(255, 255, 255)
        .Weight = xlThin
    End With

    With ws.Range("A20:H24").Borders
        .LineStyle = xlContinuous
        .Color = RGB(214, 218, 223)
        .Weight = xlThin
    End With

    ws.Range("A24").Borders(xlEdgeRight).LineStyle = xlNone
    ws.Range("B24").Borders(xlEdgeLeft).LineStyle = xlNone
    ws.Range("B24").Borders(xlEdgeRight).LineStyle = xlNone
    ws.Range("C24").Borders(xlEdgeLeft).LineStyle = xlNone

    With ws.Range("H24").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(214, 218, 223)
        .Weight = xlThin
    End With
End Sub

Private Sub AutoFitDashboardReviewRows(ByVal ws As Worksheet)
    Dim rowNumber As Long
    Dim requiredHeight As Double
    Dim wrappedLines As Long

    ws.Rows("20:23").AutoFit
    For rowNumber = 20 To 23
        wrappedLines = EstimatedWrappedStatusLines(ws.Cells(rowNumber, 3))
        requiredHeight = (wrappedLines * 11.5) + 5
        If requiredHeight < 30 Then requiredHeight = 30
        If ws.Rows(rowNumber).RowHeight < requiredHeight Then _
            ws.Rows(rowNumber).RowHeight = requiredHeight
        If ws.Rows(rowNumber).RowHeight < 30 Then ws.Rows(rowNumber).RowHeight = 30
    Next rowNumber
End Sub

Private Function EstimatedWrappedStatusLines(ByVal statusCell As Range) As Long
    Const AVERAGE_BOLD_CHARACTER_WIDTH As Double = 5.1
    Dim charactersPerLine As Long
    Dim statusLines As Variant
    Dim statusLine As Variant
    Dim lineLength As Long
    Dim lineCount As Long

    charactersPerLine = CLng(Int((statusCell.Width - 6) / AVERAGE_BOLD_CHARACTER_WIDTH))
    If charactersPerLine < 10 Then charactersPerLine = 10

    statusLines = Split(Replace$(CStr(statusCell.Value2), vbCr, vbNullString), vbLf)
    For Each statusLine In statusLines
        lineLength = Len(CStr(statusLine))
        If lineLength = 0 Then
            lineCount = lineCount + 1
        Else
            lineCount = lineCount + ((lineLength + charactersPerLine - 1) \ charactersPerLine)
        End If
    Next statusLine

    If lineCount < 1 Then lineCount = 1
    EstimatedWrappedStatusLines = lineCount
End Function

Private Sub StyleNavigationShape(ByVal navigationShape As Shape)
    navigationShape.AutoShapeType = msoShapeRoundedRectangle
    SetSubtleCornerRadius navigationShape, 0.08
    navigationShape.Fill.ForeColor.RGB = RGB(0, 43, 85)
    navigationShape.Line.ForeColor.RGB = RGB(0, 34, 68)
    navigationShape.Line.Weight = 1.25
    navigationShape.TextFrame.Characters.Font.Name = "Aptos"
    navigationShape.TextFrame.Characters.Font.Size = 10
    navigationShape.TextFrame.Characters.Font.Bold = True
    navigationShape.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    navigationShape.TextFrame.HorizontalAlignment = xlCenter
    navigationShape.TextFrame.VerticalAlignment = xlCenter
    navigationShape.Shadow.Visible = msoFalse
End Sub
