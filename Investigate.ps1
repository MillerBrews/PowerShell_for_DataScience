<#
PS4DS: Acquire (Websites)
Author: Eric K. Miller
Last updated: 28 October 2025

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
        Plot data from a supplied data object.

    .DESCRIPTION
        This function uses a data object, its fields, and optional
    parameters to create and display a PowerShell chart. The user has
    the option of selecting among several chart types.
    
    .PARAMETER DataObject
        The data object with fields to plot.

    .PARAMETER ChartType
        A ValidateSet of available chart types.
    
    .PARAMETER XData
        The x-value data from the data object.

    .PARAMETER YData
        The y-value data from the data object.

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
        Show-Chart -DataObject $StarWars_plot_fit -ChartType Point @dataParams @chartParams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$DataObject,
        [Parameter(Mandatory)]
        [ValidateSet("Bar", "Line", "Pie", "Point")][string]$ChartType,
        [Parameter()]$XData,
        [Parameter()]$YData,
        [Parameter()]$DataColor="MediumBlue",
        [Parameter()]$ChartTitleText="",
        [Parameter()]$XAxisTitleText="",
        [Parameter()]$YAxisTitleText=""
    )

    $Chart = New-Object Chart
    $ChartArea = New-Object ChartArea
    $Series = New-Object Series
    $ChartTypes = [SeriesChartType]
    $Series.ChartType = $ChartTypes::$ChartType

    $Chart.Series.Add($Series)
    $Chart.ChartAreas.Add($ChartArea)
    
    switch ($ChartType) {
    <#
        {$_ -eq "Bar"}
            {}
        {$_ -eq "Line"}
            {}
        {$_ -eq "Pie"}
            {$headers = $DataObject[0].PSObject.Properties.Name
            for ($i=0; $i -lt $headers.Length; $i++) {
                $null = $Series.Points.AddXY($labels[$i], $data[$i])
            }
            $Series["PieLabelStyle"] = "Outside"  # $Series.CustomProperties
            $Series["PieLineColor"] = "Black"  # $Series.CustomProperties
            $Series.Label = "#AXISLABEL: #VAL (#PERCENT{P0})"
            $Legend = New-Object Legend
            $Chart.Legends.Add($Legend)
            }#>
        {$_ -eq "Point"}
            {$Chart.Series['Series1'].Points.DataBindXY($XData, $YData)}
#        default
#            {$Impact_Structures[$i].'Diameter__km__approx' = $_}
    }

    $Chart.Width = 700
    $Chart.Height = 500
    $Chart.Left = 10
    $Chart.Top = 10
    $Chart.BackColor = [System.Drawing.Color]::White
    $Chart.BorderColor = 'Black'
    $Chart.BorderDashStyle = 'Solid'

    $Series.Color = [System.Drawing.Color]::$DataColor

    $ChartTitle = New-Object Title
    $ChartTitle.Text = $ChartTitleText
    $Font = New-Object System.Drawing.Font @('Lucida Console', '12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle.Font = $Font
    $Chart.Titles.Add($ChartTitle)

    $ChartArea.AxisX.Title = $XAxisTitleText
    $ChartArea.AxisY.Title = $YAxisTitleText

    $AnchorAll = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Plot Example"
    $Form.Width = 800
    $Form.Height = 600

    $Form.Controls.Add($Chart)
    $Chart.Anchor = $AnchorAll

    #$Chart.SaveImage(...)

    $Form.Add_Shown({$Form.Activate()})
    $Form.ShowDialog()
}
