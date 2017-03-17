Sub createRanking()
   Dim i As Integer
   Dim currentColor As Variant
   
   labelFontSize = 14
   labelWidth = 200
   
   bgColorBrightness = 0.5
   
   
   fixRed = 0
   loRed = 150
   hiRed = 200
   
   fixGreen = 0
   loGreen = 150
   hiGreen = 200
   
   fixBlue = 0
   loBlue = 150
   hiBlue = 200
   
   drawLines = 0

   randomizeColors = 1
   
   maxI = ActiveSheet.ChartObjects(1).Chart.SeriesCollection.Count
   
   With ActiveSheet.ChartObjects(1).Chart
      For i = 1 To .SeriesCollection.Count
        .SeriesCollection
         With .SeriesCollection(i)
         
            Randomize
            
            If randomizeColors = 1 Then
                randomRed = Int((hiRed - loRed + 1) * Rnd() + loRed)
            Else
                randomRed = Int(loRed + ((hiRed - loRed) * i / maxI))
            End If
            If fixRed = 0 Then
                red = randomRed
            Else
                red = fixRed
            End If
          
            If randomizeColors = 1 Then
                randomGreen = Int((hiGreen - loGreen + 1) * Rnd() + loGreen)
            Else
                randomColor = Int(loGreen + ((hiGreen - loGreen) * i / maxI))
            End If
            
            If fixGreen = 0 Then
                green = randomGreen
            Else
                green = fixGreen
            End If
            
            If randomizeColors = 1 Then
                randomBlue = Int((hiBlue - loBlue + 1) * Rnd() + loBlue)
            Else
                randomBlue = Int(loBlue + ((hiBlue - loBlue) * i / maxI))
            End If
            If fixBlue = 0 Then
                blue = randomBlue
            Else
                blue = fixBlue
            End If
            
            
            
            currentColor = RGB(red, green, blue)
            
            .ApplyDataLabels
            .DataLabels.ShowCategoryName = False
            .DataLabels.ShowSeriesName = True
            .DataLabels.Position = xlLabelPositionCenter
            .DataLabels.ShowValue = False
            .DataLabels.Format.Fill.ForeColor.RGB = currentColor
            .DataLabels.Format.Fill.ForeColor.Brightness = bgColorBrightness
            .DataLabels.Format.Fill.Transparency = 0
            .DataLabels.Format.Fill.Solid
          
            .DataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .DataLabels.Format.TextFrame2.TextRange.Font.Size = labelFontSize
            
            Values_Array = .Values
            For j = LBound(Values_Array, 1) To UBound(Values_Array, 1)
               .Points(j).DataLabel.Width = labelWidth
            Next
  
            If drawLines = 1 Then
                .Format.Line.Visible = msoTrue
                .Format.Line.ForeColor.RGB = currentColor
                .Format.Line.ForeColor.TintAndShade = 0
                .Format.Line.ForeColor.Brightness = bgColorBrightness
                .Format.Line.Transparency = 0
                .Format.Line.Weight = 2
            Else
                .Format.Line.Visible = msoFalse
            End If
            
            
         End With
         
      Next i
   End With
End Sub

