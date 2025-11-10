<#
PS4DS: Acquire (Websites)
Author: Eric K. Miller
Last updated: 10 November 2025

This script contains PowerShell code for plotting data.
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
    optional parameters to create and display a PowerShell chart. The
    user has the option of selecting among several chart types.
    
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

    .PARAMETER ChartTitleText
        Enables the user to set title text for the chart.

    .PARAMETER XAxisTitleText
        Enables the user to set title text for the x-axis.
    
    .PARAMETER YAxisTitleText
        Enables the user to set title text for the y-axis.

    .EXAMPLE
        $dataParams = @{XData = $StarWars_plot_fit.'height'
                YData = $StarWars_plot_fit.'mass'
        }
        $chartParams = @{ChartTitleText = "Star Wars Height vs. Mass"
            XAxisTitleText = "Height (cm)"
            YAxisTitleText = "Mass (kg)"
        }
        Show-Chart -ChartType Point @dataParams @chartParams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
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
    $Chart.Left = 10
    $Chart.Top = 10
    $Chart.BackColor = [System.Drawing.Color]::White
    $Chart.BorderColor = 'Black'
    $Chart.BorderDashStyle = 'Solid'
    $ChartArea = New-Object ChartArea
    $Chart.ChartAreas.Add($ChartArea)

    # Create form to host the chart
    $Form = New-Object Windows.Forms.Form
    #$Form.Text = 'Plot Example'
    $Form.Width = 640
    $Form.Height = 480
    $Form.Controls.Add($Chart)
    
    # Create the data series
    switch ($ChartType) {
        'Bar' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType

            $Series.Name = 'BarPlotSeries'
            $Series.Points.DataBindXY($XData, $YData)
            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $ChartArea.AxisX.Interval = 1
            $Chart.Series.Add($Series)
        }
        'Column' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType

            $Series.Name = 'ColumnPlotSeries'
            $Series.Points.DataBindXY($XData, $YData)
            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $ChartArea.AxisX.Interval = 1
            $Chart.Series.Add($Series)
        }
        'Line' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType
            
            $XData_bounds = $XData | Measure-Object -Minimum -Maximum
            $x_pts = $XData_bounds.Minimum..$XData_bounds.Maximum
            $y_pts = $x_pts | ForEach-Object {$Theta[0] * $_ + $Theta[1]}

            $Series.Name = 'LinePlotSeries'
            $Series.Points.DataBindXY($x_pts, $y_pts)
            #$Series.Color = [System.Drawing.Color]::Thistle
            $Chart.Series.Add($Series)
        }
        'Pie' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType
            
            for ($i = 0; $i -lt $XData.Count; $i++) {
                $null = $Series.Points.AddXY($XData[$i], $YData[$i])
            }
            $Series['PieLabelStyle'] = 'Outside'  # $Series.CustomProperties
            $Series['PieLineColor'] = 'Black'  # $Series.CustomProperties
            $Series.Label = "#AXISLABEL: #VAL (#PERCENT{P0})"
            $Chart.Series.Add($Series)
            #$Legend = New-Object Legend
            #$Chart.Legends.Add($Legend)
        }
        'Point' {
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::$ChartType
            
            $Series.Name = 'PointPlotSeries'
            $Series.Points.DataBindXY($XData, $YData)
            #$Series.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($Series)
        }
        'PointsAndLine' {
            # Add points series
            $PointsSeries = New-Object Series
            $ChartTypes = [SeriesChartType]
            $PointsSeries.ChartType = $ChartTypes::Point

            $PointsSeries.Name = 'PointsPlotSeries'
            $PointsSeries.Points.DataBindXY($XData, $YData)
            #$PointsSeries.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($PointsSeries)

            # Add line series
            $LineSeries = New-Object Series
            $LineSeries.ChartType = $ChartTypes::Line

            $XData_bounds = $XData | Measure-Object -Minimum -Maximum
            $x_pts = $XData_bounds.Minimum..$XData_bounds.Maximum
            $y_pts = $x_pts | ForEach-Object {$Theta[0] * $_ + $Theta[1]}

            $LineSeries.Name = 'LinePlotSeries'
            $LineSeries.Points.DataBindXY($x_pts, $y_pts)
            #$LineSeries.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($LineSeries)
        }
        'BoxPlot' {
            $rawSeries = New-Object Series
            $ChartTypes = [SeriesChartType]
            $rawSeries.ChartType = $ChartTypes::Point

            $rawSeries.Name = 'RawData'
            $rawSeries.IsVisibleInLegend = $false
            $rawSeries.IsValueShownAsLabel = $false
            
            foreach ($value in $XData) {
                $point = New-Object DataPoint
                $point.XValue = $groupX
                $point.YValues = @($value)
                $rawSeries.Points.Add($point)
            }
            $Chart.Series.Add($rawSeries)

            # BoxPlot series
            $boxSeries = New-Object Series
            $ChartTypes = [SeriesChartType]
            $boxSeries.ChartType = $ChartTypes::BoxPlot

            $boxSeries.Name = 'BoxPlot'
            $boxSeries['BoxPlotSeries'] = 'RawData'
            $boxSeries['BoxPlotShowMedian'] = $true
            $boxSeries['BoxPlotShowUnusualValues'] = $true
            $boxSeries['BoxPlotShowExtremeValues'] = $true
            $boxSeries['BoxPlotWhiskerPercentile'] = '10'

            # Add a dummy point to trigger boxplot rendering
            [void]$boxSeries.Points.AddXY(1, 0)
            $Chart.Series.Add($boxSeries)
        }
        'Histogram' {
            # Create histogram with n buckets
            $histogram = New-Object MathNet.Numerics.Statistics.Histogram($XData, 10)

            # Create series for histogram bars
            $Series = New-Object Series
            $ChartTypes = [SeriesChartType]
            $Series.ChartType = $ChartTypes::'Column'
            $Series.Name = 'Histogram'

            $Series.Color = [System.Drawing.Color]::$DataColor
            $Series.BorderWidth = 1

            # Add buckets to chart
            for ($i = 0; $i -lt $histogram.BucketCount; $i++) {
                $bucket = $histogram[$i]
                $label = "{0:N1}-{1:N1}" -f $bucket.LowerBound, $bucket.UpperBound
                [void]$Series.Points.AddXY($label, $bucket.Count)
            }
            $Chart.Series.Add($Series)
        }
#        default
#            {$Impact_Structures[$i].'Diameter__km__approx' = $_}
    }
    
    $ChartTitle = New-Object Title
    $ChartTitle.Text = $ChartTitleText
    $Font = New-Object System.Drawing.Font @('Lucida Console', '12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle.Font = $Font
    $Chart.Titles.Add($ChartTitle)

    $ChartArea.AxisX.Title = $XAxisTitleText
    $ChartArea.AxisY.Title = $YAxisTitleText

    $AnchorAll = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

    $Chart.Anchor = $AnchorAll
    $Chart.Dock = "Fill"

    #$Chart.SaveImage(...)

    $Form.Add_Shown({$Form.Activate()})
    $Form.ShowDialog()
}
