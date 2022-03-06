Private Sub CommandButton1_Click()

Dim my_chart As ChartObject
Dim var As range
Dim srs As Series
Dim MaxNumber As Double
Dim MinNumber As Double
Dim SeriesNumber As Integer

'Data range for chart
Set var = ActiveSheet.range(range("O3").Value)

'Draw chart
Set my_chart = ActiveSheet.ChartObjects.Add( _
    Left:=ActiveCell.Left, _
    Width:=450, _
    Top:=ActiveCell.Top, _
    Height:=250)

'Supply data to chart
my_chart.Chart.SetSourceData Source:=var

'Determine the chart type
my_chart.Chart.ChartType = xlLine

'Chart title
my_chart.Chart.HasTitle = True
my_chart.Chart.ChartTitle.Text = range("O4").Value
my_chart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
my_chart.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue

'Remove gridlines
my_chart.Chart.Axes(xlValue).MajorGridlines.Delete

'Date formatting
my_chart.Chart.Axes(xlCategory).MajorUnit = 1
my_chart.Chart.Axes(xlCategory).MajorUnitScale = xlYears
my_chart.Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"
my_chart.Chart.Axes(xlCategory).TickLabels.Font.Size = 9
my_chart.Chart.Axes(xlCategory).MajorTickMark = xlNone

'Remove value tickmarks and make axis white
my_chart.Chart.Axes(xlValue).MajorTickMark = xlNone
my_chart.Chart.Axes(xlValue).TickLabels.Font.Size = 9
my_chart.Chart.Axes(xlValue).Border.Color = vbWhite

'Legend at bottom
my_chart.Chart.SetElement (msoElementLegendBottom)
my_chart.Chart.Legend.Font.Size = 9

'Auto-adjust axes
    For Each srs In my_chart.Chart.SeriesCollection
    
        'Determine max value in series
            MaxNumber = Application.WorksheetFunction.Max(srs.Values)
                        
        'Determine min value in series
            MinNumber = Application.WorksheetFunction.Min(srs.Values)
            
        'Determine number of values in series
            SeriesNumber = Application.WorksheetFunction.Count(srs.Values)
            
        'Adjust x-axis scale if SeriesNumber is less than 366 (i.e. when there's a year or less of data)
            If SeriesNumber < 366 Then
                my_chart.Chart.Axes(xlCategory).MajorUnitScale = xlMonths
            End If
            
        'Rescale y-axis
            my_chart.Chart.Axes(xlValue).MinimumScale = Application.WorksheetFunction.Floor(MinNumber, 10)
            my_chart.Chart.Axes(xlValue).MaximumScale = Application.WorksheetFunction.Ceiling(MaxNumber, 10)
                        
    Next srs
                

End Sub
