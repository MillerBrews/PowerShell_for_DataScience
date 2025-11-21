using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

$StarWars_NoJabba = $StarWars | ?{$_.mass -lt 200 -and $_.mass -ne $null}

$numeric_fields = $StarWars[0].PSObject.Properties |
    Where-Object {$_.Value -is [double] -or $_.Value -is [decimal]} |
    ForEach-Object {$_.Name}

#region Boilerplate
$Chart = New-Object Chart
$Chart.Width = 700
$Chart.Height = 500
$ChartArea = New-Object ChartArea
$ChartArea.Name = 'ChartArea'
$Chart.ChartAreas.Add($ChartArea)

# Create form to host the chart
$Form = New-Object Windows.Forms.Form
$Form.Width = 640
$Form.Height = 480
$Form.Controls.Add($Chart)
#endregion

$Form.Text = 'BoxPlot'

$data_colors = @(
    [System.Drawing.Color]::LightBlue,
    [System.Drawing.Color]::LightGoldenrodYellow,
    [System.Drawing.Color]::LightGreen,
    [System.Drawing.Color]::LightPink,
    [System.Drawing.Color]::LightSteelBlue,
    [System.Drawing.Color]::LightYellow
)

$fields_index = 0
foreach ($field in $numeric_fields) {
    # Extract values for each field
    $field_values = $StarWars_NoJabba | ForEach-Object {$_.$field}

    # Point data series
    $Series = New-Object Series
    $Series.Name = "PointData_$fields_index"
    $Series.ChartType = [SeriesChartType]::Point
    $Series.ChartArea = $ChartArea.Name
    $Series.MarkerStyle = [MarkerStyle]::Circle
    $Series.MarkerSize = 6
    $Series.Color = 'Black'
    [double[]]$field_values | ForEach-Object {[void]$Series.Points.AddXY($fields_index+1, $_)}
    $Chart.Series.Add($Series)

    # BoxPlot series
    $BoxSeries = New-Object Series
    $BoxSeries.Name = "BoxPlot_$fields_index"
    $BoxSeries.ChartType = [SeriesChartType]::BoxPlot
    $BoxSeries.ChartArea = $ChartArea.Name
    
    $BoxSeries['BoxPlotSeries'] = $Series.Name
    <#
    $BoxSeries['BoxPlotShowAverage'] = 'true'
    $BoxSeries['BoxPlotShowMedian'] = 'true'
    $BoxSeries['BoxPlotShowInnerPoints'] = 'true'
    $BoxSeries['BoxPlotShowUnusualValues'] = 'true'
    $BoxSeries['BoxPlotShowExtremeValues'] = 'true'
    $BoxSeries['BoxPlotShowExtremeValues'] = 'true'
    $BoxSeries['BoxPlotWhiskerPercentile'] = '10'
    $BoxSeries.Color = $data_colors[$fields_index % $data_colors.Length]
    $BoxSeries.BorderColor = 'LightGray'
    $BoxSeries.BorderWidth = 2
    #>
    [void]$BoxSeries.Points.AddXY(1, 0)
    [void]$BoxSeries.Points.AddXY($fields_index + 1, 0)  # not working
    $Chart.Series.Add($BoxSeries)
        
    $Label = New-Object CustomLabel
    $Label.FromPosition = $fields_index + 1 - 0.5
    $Label.ToPosition = $fields_index + 1 + 0.5
    $Label.Text = $field
    $ChartArea.AxisX.CustomLabels.Add($Label)

    $fields_index++
}
$ChartArea.AxisX.Interval = 1
$ChartArea.AxisX.Minimum = 0

$Chart.Dock = 'Fill'
$Form.Controls.Add($Chart)
$Form.ShowDialog()