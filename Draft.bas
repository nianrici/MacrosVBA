Sub Draft()
    Dim oTable As Table
    Dim var As String
    Dim textVar As String
    Dim imatge As InlineShape
    Dim lSection As Section
    Dim contador As Integer
    
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
    For Each lSection In ActiveDocument.Sections
        lSection.Range.Select
        Dim strWMName As String
        strWMName = lSection.Index
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Selection.HeaderFooter.Shapes.AddTextEffect(msoTextEffect1, _
        "DRAFT", "Arial", 1, False, False, 0, 0).Select
        With Selection.ShapeRange
            .Name = strWMName
            .TextEffect.NormalizedHeight = False
            .Line.Visible = False
            With .Fill
                .Visible = True
                .Solid
                .ForeColor.RGB = Gray
                .Transparency = 0.9
            End With
            .Rotation = 315
            .LockAspectRatio = True
            .Height = InchesToPoints(2.42)
            .Width = InchesToPoints(6.04)
            With .WrapFormat
                .AllowOverlap = True
                .Side = wdWrapNone
                .Type = 3
            End With
            .RelativeHorizontalPosition = wdRelativeVerticalPositionMargin
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            .Top = wdShapeCenter
        End With
        
    Next lSection
    
    With ActiveDocument.Sections(1)
        .Headers(wdHeaderFooterFirstPage).Shapes.AddTextEffect(msoTextEffect1, _
        "DRAFT", "Arial", 1, False, False, 0, 0).Select
        With Selection.ShapeRange
            .TextEffect.NormalizedHeight = False
            .Line.Visible = False
            With .Fill
                .Visible = True
                .Solid
                .ForeColor.RGB = Gray
                .Transparency = 0.9
            End With
            .Rotation = 315
            .LockAspectRatio = True
            .Height = InchesToPoints(2.42)
            .Width = InchesToPoints(6.04)
            With .WrapFormat
                .AllowOverlap = True
                .Side = wdWrapNone
                .Type = 3
            End With
            .RelativeHorizontalPosition = wdRelativeVerticalPositionMargin
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            .Top = wdShapeCenter
        End With
    End With
End Sub