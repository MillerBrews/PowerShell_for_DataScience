using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

$StarWars_NoJabba = $StarWars | ?{$_.mass -lt 200 -and $_.mass -ne $null}

$numeric_fields = $StarWars[0].PSObject.Properties |
    Where-Object {$_.Value -is [double] -or $_.Value -is [decimal]} |
    ForEach-Object {$_.Name}

#region Boilerplate
# Form to hold the chart
$Form = New-Object Windows.Forms.Form
$Form.Width = 640
$Form.Height = 480
$Form.Controls.Add($Chart)

# Chart to hold the ChartArea and Series
$Chart = New-Object Chart
$Chart.Width = 700
$Chart.Height = 500
$ChartArea = New-Object ChartArea
$ChartArea.Name = 'ChartArea'
$Chart.ChartAreas.Add($ChartArea)

$BoxSeries = New-Object Series
$BoxSeries.Name = "BoxPlot_Data"
$BoxSeries.ChartType = [SeriesChartType]::BoxPlot
$BoxSeries.ChartArea = $ChartArea.Name
#$BoxSeries['BoxPlotPercentile'] = '25'
$BoxSeries['BoxPlotShowAverage'] = 'true'
$BoxSeries['BoxPlotShowMedian'] = 'true'
$BoxSeries['BoxPlotShowInnerPoints'] = 'false'
$BoxSeries['BoxPlotShowUnusualValues'] = 'true'
$BoxSeries['BoxPlotShowExtremeValues'] = 'true'
$BoxSeries['BoxPlotWhiskerPercentile'] = '25'
$BoxSeries['PointWidth'] = '0.5'
$BoxSeries.Color = $data_colors[$fields_index % $data_colors.Length]
$BoxSeries.BorderColor = 'LightGray'
$BoxSeries.BorderWidth = 2
$Chart.Series.Add($BoxSeries)
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
    $PointSeries = New-Object Series
    $PointSeries.Name = "PointData_$fields_index"
    $PointSeries.ChartType = [SeriesChartType]::Point
    $PointSeries.ChartArea = $ChartArea.Name
    $PointSeries.MarkerStyle = [MarkerStyle]::Circle
    $PointSeries.MarkerSize = 6
    $PointSeries.Color = 'Black'
    [double[]]$field_values | ForEach-Object {[void]$PointSeries.Points.AddXY($fields_index+1, $_)}
    $Chart.Series.Add($PointSeries)

    $field_values_sorted = $field_values | Sort-Object
    $min = $field_values_sorted[0]
    $max = $field_values_sorted[-1]
    $median = if ($field_values_sorted.Length % 2 -eq 0) {
            ($field_values_sorted[$field_values_sorted.Length / 2] +
            $field_values_sorted[$field_values_sorted.Length / 2 - 1]) / 2
        }
        else { $field_values_sorted[[Math]::Floor($field_values_sorted.Length / 2)] }
    $q1_indx = [Math]::Floor(($field_values_sorted.count-1) * 0.25)
    $q3_indx = [Math]::Floor(($field_values_sorted.count-1) * 0.75)
    $q1 = $field_values_sorted[$q1_indx]
    $q3 = $field_values_sorted[$q3_indx]
    $avg = ($field_values | Measure-Object -Average).Average

    # BoxPlot series
    $BoxPoints = New-Object DataPoint
    $BoxPoints.SetValueXY($fields_index + 1, @($min, $q1, $median, $q3, $max, $avg))
    $BoxSeries.Points.Add($BoxPoints)
    #
    <#
    $BoxSeries = New-Object Series
    $BoxSeries.Name = "BoxPlot_$fields_index"
    $BoxSeries.ChartType = [SeriesChartType]::BoxPlot
    $BoxSeries.ChartArea = $ChartArea.Name
    $BoxSeries["BoxPlotSeries"] = $PointSeries.Name
    $BoxSeries["BoxPlotShowMedian"] = "true"
    $BoxSeries["BoxPlotShowUnusualValues"] = "true"
    $BoxSeries["BoxPlotShowExtremeValues"] = "true"
    [void]$BoxSeries.Points.AddXY($fields_index + 1, 0)
    #[double[]]$field_values | ForEach-Object {[void]$BoxSeries.Points.AddXY($fields_index + 1, $_)}
    $Chart.Series.Add($BoxSeries)
    #>
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