<#
PS4DS: Investigate (Plotting)
Author: Eric K. Miller
Last updated: 20 November 2025

This script contains PowerShell code for plotting data. It is part of
the PowerShell_for_DataScience module, so assumes the Math.NET Numerics
DLL is loaded, as this is integral to the histogram plotting.
#>

#========================
# Investigate functions
#========================

using namespace System.Windows.Forms.DataVisualization.Charting
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

function Show-Chart {
    <#
    .SYNOPSIS
        Plot data from supplied data series.

    .DESCRIPTION
        This function uses fields from a data object (arrays), and
    optional parameters to create and display a PowerShell chart.
    The user has the option to select among several chart types.
    
    .PARAMETER DataObject
        The object to use to plot data.

    .PARAMETER ChartType
        A ValidateSet of available chart types.
    
    .PARAMETER XData
        The x-value data from the data object.

    .PARAMETER YData
        The y-value data from the data object.

    .PARAMETER Theta
        The parameters for a line, i.e., @(slope, y_intercept).

    .PARAMETER DataColor
        Enables the user to set the primary chart color.
    (Default is 'SteelBlue').

    .PARAMETER ChartTitleText
        Enables the user to set title text for the chart.

    .PARAMETER XAxisTitleText
        Enables the user to set text for the x-axis.
    
    .PARAMETER YAxisTitleText
        Enables the user to set text for the y-axis.

    .EXAMPLE
        $dataParams = @{XData = $StarWars.'height'
                        YData = $StarWars.'mass'}
        $chartParams = @{ChartTitleText = "Star Wars Height vs. Mass"
            XAxisTitleText = "Height (cm)"
            YAxisTitleText = "Mass (kg)"}
        Show-Chart -ChartType Point @dataParams @chartParams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter()]$DataObject,
        [Parameter(Mandatory)]
        [ValidateSet('Bar', 'Column', 'Line', 'Pie', 'Point', 'PointsAndLine', 'BoxPlot', 'Histogram')]
        [string]$ChartType,
        [Parameter()]$XData=$null,
        [Parameter()]$YData=$null,
        [Parameter()]$Theta=$null,
        [Parameter()]$DataColor='SteelBlue',
        [Parameter()]$ChartTitleText='',
        [Parameter()]$XAxisTitleText='',
        [Parameter()]$YAxisTitleText=''
    )

    # Create chart object and set its properties
    $Chart = New-Object Chart
    $Chart.Width = 700
    $Chart.Height = 500
    $ChartArea = New-Object ChartArea
    $Chart.ChartAreas.Add($ChartArea)

    # Create form to host the chart
    $Form = New-Object Windows.Forms.Form
    $Form.Width = 640
    $Form.Height = 480
    $Form.Controls.Add($Chart)
    
    # Create the data series and their properties for charting
    switch ($ChartType) {
        'Bar' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType

            $Series.Points.DataBindXY($XData, $YData)
            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $ChartArea.AxisX.Interval = 1
            $Chart.Series.Add($Series)

            $Form.Text = 'Bar Plot'
        }
        'Column' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType

            $Series.Points.DataBindXY($XData, $YData)
            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $ChartArea.AxisX.Interval = 1
            $Chart.Series.Add($Series)

            $Form.Text = 'Column Plot'
        }
        'Line' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType
            
            $XData_bounds = $XData | Measure-Object -Minimum -Maximum
            $x_pts = $XData_bounds.Minimum..$XData_bounds.Maximum
            $y_pts = $x_pts | ForEach-Object {$Theta[0] * $_ + $Theta[1]}

            $Series.Points.DataBindXY($x_pts, $y_pts)
            $Series.Color = [System.Drawing.Color]::$DataColor
            $Series.BorderWidth = 3
            
            # Create a text annotation
            $annotation = New-Object TextAnnotation
            $annotation.Text = "y = $([Math]::Round($Theta[0],2))*x + $([Math]::Round($Theta[1],2))"
            $annotation.Font = New-Object System.Drawing.Font('Lucida Console', 10, [System.Drawing.FontStyle]::Bold)
            $annotation.ForeColor = [System.Drawing.Color]::$DataColor
            $annotation.AxisX = $Chart.ChartAreas[0].AxisX
            $annotation.AxisY = $Chart.ChartAreas[0].AxisY
            $annotation_Xpos = 1.75 * ($XData | Measure-Object -Minimum).Minimum
            $annotation.AnchorX = $annotation_Xpos
            $annotation.AnchorY = 1.5 * ($Theta[0] * $annotation_Xpos + $Theta[1])
            #$annotation.IsSizeAlwaysRelative = $false
            
            $Chart.Series.Add($Series)
            $Chart.Annotations.Add($annotation)

            $Form.Text = 'Line Plot'
        }
        'Pie' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType

            for ($i = 0; $i -lt $XData.Count; $i++) {
                [void]$Series.Points.AddXY($XData[$i], $YData[$i])
            }
            
            $Series['PieLabelStyle'] = 'Outside'  # $Series.CustomProperties
            $Series['PieLineColor'] = 'Gray'  # $Series.CustomProperties
            $Series.Label = "#AXISLABEL: #VAL (#PERCENT{P0})"
            $Chart.Series.Add($Series)
            
            $Form.Text = 'Pie Plot'
        }
        'Point' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType
            
            $Series.Points.DataBindXY($XData, $YData)
            $Series.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($Series)

            $Form.Text = 'Scatter Plot'
        }
        'PointsAndLine' {
            # Add points series -------------------------------
            $PointsSeries = New-Object Series
            $ChartTypes = [SeriesChartType]
            $PointsSeries.ChartType = $ChartTypes::Point
            $PointsSeries.LegendText = 'Raw data'

            $PointsSeries.Points.DataBindXY($XData, $YData)
            $PointsSeries.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($PointsSeries)

            # Add line series  --------------------------------
            $LineSeries = New-Object Series
            $LineSeries.ChartType = $ChartTypes::Line
            $LineSeries.LegendText = 'Regression line'

            $XData_bounds = $XData | Measure-Object -Minimum -Maximum
            $x_pts = $XData_bounds.Minimum..$XData_bounds.Maximum
            $y_pts = $x_pts | ForEach-Object {$Theta[0] * $_ + $Theta[1]}

            $LineSeries.Points.DataBindXY($x_pts, $y_pts)
            $lineColor = 'Goldenrod'
            $LineSeries.Color = [System.Drawing.Color]::$lineColor
            $LineSeries.BorderWidth = 3
            
            # Create a text annotation showing the line's equation
            $annotation = New-Object TextAnnotation
            $annotation.Text = "y = $([Math]::Round($Theta[0],2))*x + $([Math]::Round($Theta[1],2))"
            $annotation.Font = New-Object System.Drawing.Font('Lucida Console', 10, [System.Drawing.FontStyle]::Bold)
            $annotation.ForeColor = [System.Drawing.Color]::$lineColor
            $annotation.AxisX = $Chart.ChartAreas[0].AxisX
            $annotation.AxisY = $Chart.ChartAreas[0].AxisY
            $annotation_Xpos = 1.75 * ($XData | Measure-Object -Minimum).Minimum
            $annotation.AnchorX = $annotation_Xpos
            $annotation.AnchorY = 1.5 * ($Theta[0] * $annotation_Xpos + $Theta[1])
            
            $Chart.Series.Add($LineSeries)
            $Chart.Annotations.Add($annotation)

            # Create legend
            $legend = New-Object Legend
            $legend.Docking = 'Top'
            $legend.Alignment = 'Center'
            $Chart.Legends.Add($legend)

            $Form.Text = 'Point and Line Plot'
        }
        'BoxPlot' {
            # Loop through each numeric column
            for ($i = 0; $i -lt $XData[0].PSObject.Properties.Name.Count; $i++) {
                # Add points series -------------------------------
                # ...
                # Add BoxPlot series ------------------------------
                $col = $XData[0].PSObject.Properties.Name[$i]
                $xVal = $i + 1

                # Raw data series
                $rawSeries = New-Object Series
                $rawSeries.Name = "Raw_$col"
                $rawSeries.ChartType = 'Point'
                $rawSeries.IsVisibleInLegend = $false

                foreach ($row in $data) {
                    $val = $row.$col
                    if ($val -is [double] -or $val -is [int]) {
                        $point = New-Object DataPoint
                        $point.XValue = $xVal
                        $point.YValues = @($val)
                        $rawSeries.Points.Add($point)
                    }
                }
                $Chart.Series.Add($rawSeries)

                # BoxPlot series
                $boxSeries = New-Object Series
                $boxSeries.Name = "Box_$col"
                $boxSeries.ChartType = 'BoxPlot'
                $boxSeries['BoxPlotSeries'] = $rawSeries.Name
                $boxSeries['BoxPlotShowMedian'] = "true"
                $boxSeries['BoxPlotShowUnusualValues'] = "true"
                $boxSeries['BoxPlotShowExtremeValues'] = "true"
                $boxSeries['BoxPlotWhiskerPercentile'] = "10"
                $boxSeries.BorderColor = 'Black'
                $boxSeries.BorderWidth = 2

                # Trigger rendering
                [void]$boxSeries.Points.AddXY($xVal, 0)
                $Chart.Series.Add($boxSeries)

                # Label
                $label = New-Object CustomLabel
                $label.FromPosition = $xVal - 0.5
                $label.ToPosition = $xVal + 0.5
                $label.Text = $col
                $ChartArea.AxisX.CustomLabels.Add($label)
               }

            # Color outliers in red
            foreach ($series in $Chart.Series) {
                if ($series.ChartType -eq 'BoxPlot') {
                    foreach ($pt in $series.Points) {
                        if ($pt.MarkerStyle -eq 'Circle') {
                            $pt.Color = 'Red'
                        }
                    }
                }
            }
            $Chart.ChartAreas[0].AxisX.Interval = 1
            $Chart.ChartAreas[0].AxisX.IsLabelAutoFit = $false
            $ChartArea.AxisX.Minimum = 0

            $Chart.Series.Add($boxSeries)

            $Form.Text = 'BoxPlots by Numeric Column'
            #>
        }
        'Histogram' {
            # Create histogram with n buckets
            $n = 10
            $histogram = New-Object MathNet.Numerics.Statistics.Histogram($XData, $n)

            # Create series for histogram bars
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::'Column'

            $Series.Color = [System.Drawing.Color]::$DataColor
            $Series.IsValueShownAsLabel = $true
            
            $ChartArea.AxisX.Interval = 1
            # Remove grid lines
            $ChartArea.AxisX.MajorGrid.LineWidth = 0
            $ChartArea.AxisY.MajorGrid.LineWidth = 0

            # Add buckets to chart
            for ($i = 0; $i -lt $histogram.BucketCount; $i++) {
                $bucket = $histogram[$i]
                $label = "{0:N1}-{1:N1}" -f $bucket.LowerBound, $bucket.UpperBound
                [void]$Series.Points.AddXY($label, $bucket.Count)
            }
            $Chart.Series.Add($Series)
            
            $Form.Text = 'Histogram'
        }
    }
    
    $ChartTitle = New-Object Title
    $ChartTitle.Text = $ChartTitleText
    $Font = New-Object System.Drawing.Font('Lucida Console', 12, [System.Drawing.FontStyle]::Bold)
    $ChartTitle.Font = $Font
    $Chart.Titles.Add($ChartTitle)

    $ChartArea.AxisX.Title = $XAxisTitleText
    $ChartArea.AxisY.Title = $YAxisTitleText
    
    $Chart.Dock = 'Fill'  # chart adjusts to fit the entire container when the Form is resized

    #$Chart.SaveImage(...)

    $Form.Add_Shown({$Form.Activate()})  # ensures the Form gets focus
    $Form.ShowDialog()  # shows the Form
}