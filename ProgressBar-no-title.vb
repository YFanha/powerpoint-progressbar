Sub AutoSections()
     
    Dim intSlide As Integer
    Dim strNotes As String
    Dim tabSectionNames() As String
    Dim tabSectionSlides() As Integer
    Dim sec As Integer
    Dim secNumber As Integer
    sec = 1
    Dim width As Integer
    Dim barWidth As Integer
    Dim normalColor
    Dim emphColor
    Dim BackgroundColor
    Dim j As Integer
    Dim visibleSlides As Integer

    visibleSlides = 0

    'Parameters COLORS -----------------------------------------------------------------------------------
    normalColor = RGB(233, 113, 50)
    emphColor = RGB(208, 58, 42) 'here
    BackgroundColor = RGB(255, 181, 143) 'here
    
    ' Set width to 90% of the slide width
    width = ActivePresentation.PageSetup.SlideWidth * 0.9
    Dim leftPos As Single
    leftPos = (ActivePresentation.PageSetup.SlideWidth - width) / 2 ' Centered position
    'leftPos = 75
    topPos = 20


    With ActivePresentation
        ' Loop to identify sections based on slide notes
        For intSlide = 1 To .Slides.Count
            strNotes = .Slides(intSlide).NotesPage. _
                Shapes.Placeholders(2).TextFrame.TextRange.Lines(1).Text
            strNotes = Replace(strNotes, vbLf, "")
            strNotes = Replace(strNotes, vbCr, "")

            If InStr(strNotes, "Section:") = 1 Then
                ReDim Preserve tabSectionNames(sec + 1)
                ReDim Preserve tabSectionSlides(sec + 1)
                tabSectionNames(sec) = Mid(strNotes, 9)
                tabSectionSlides(sec) = intSlide
                sec = sec + 1
            End If
            If Not .Slides(intSlide).SlideShowTransition.Hidden Then
                visibleSlides = visibleSlides + 1
            End If
        Next intSlide
        
        secNumber = sec - 1

        ' Loop through slides to draw progress bar
        For intSlide = 1 To .Slides.Count
            j = 1
            While j <= .Slides(intSlide).Shapes.Count
                ' Delete existing progress bar shapes
                If .Slides(intSlide).Shapes(j).Name = "MyBar1" Or _
                   .Slides(intSlide).Shapes(j).Name = "MyBar2" Then
                    .Slides(intSlide).Shapes(j).Delete
                Else
                    j = j + 1
                End If
            Wend
            
            ' Draw background bar
            With .Slides(intSlide).Shapes.AddShape(Type:=msoShapeRectangle, _
                    Left:=leftPos, Top:=topPos, width:=width, Height:=3)
                .Fill.BackColor.RGB = BackgroundColor
                .Fill.ForeColor.RGB = BackgroundColor
                .Line.Visible = False
                .Name = "MyBar1"
            End With
            
            ' Calculate and draw active progress bar
            barWidth = CInt(width * (intSlide - 1#) / (visibleSlides - 1#))
            With .Slides(intSlide).Shapes.AddShape(Type:=msoShapeRectangle, _
                    Left:=leftPos, Top:=topPos, width:=barWidth, Height:=3)
                .Fill.BackColor.RGB = emphColor
                .Fill.ForeColor.RGB = emphColor
                .Line.Visible = False
                .Name = "MyBar2"
            End With
        Next intSlide
    End With

End Sub




