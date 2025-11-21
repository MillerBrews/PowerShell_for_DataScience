using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

$StarWars_NoJabba = $StarWars | ?{$_.mass -lt 200}

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

function Get-BoxPlotStatistics {
    param(
        [Parameter()][double[]]$DataArray
    )
    
    if ($DataArray.Count -eq 0) {return @(0,0,0,0,0,0)}
    
    $sorted_data = $DataArray | Sort-Object
    $min = $sorted_data[0]
    $max = $sorted_data[-1]

    #Median
    $median = if ($sorted_data.Lengths % 2 -eq 0) {
        ($sorted_data[$sorted_data.Length / 2] + $sorted_data[$sorted_data.Length / 2 - 1]) / 2
        }
        else { $sorted_data[[Math]::Floor($sorted_data.Length / 2)] }

    #Q1
    $lower_half = $sorted_data[0..([Math]::Floor($sorted_data.Length / 2) - 1)]
    if ($lower_half.Count -eq 0) {
        $q1 = $min
    }
    else {
        $q1 = if ($lower_half.Length % 2 -eq 0) {
            ($lower_half[$lower_half.Length / 2] + $lower_half[$lower_half.Length / 2 - 1]) / 2
        }
        else { $lower_half[[Math]::Floor($lower_half.Length / 2)] }
    }

    #Q3
    $upper_half = $sorted_data[[Math]::Ceiling($sorted_data.Length / 2)..($sorted_data.Length - 1)]
    if ($upper_half.Count -eq 0) {
        $q3 = $max
    }
    else {
        $q3 = if ($upper_half.Length % 2 -eq 0) {
            ($upper_half[$upper_half.Length / 2] + $upper_half[$upper_half.Length / 2 - 1]) / 2
        }
        else { $upper_half[[Math]::Floor($upper_half.Length / 2)] }
    }

    $avg = ($DataArray | Measure-Object -Average).Average

    return @{
        Min = $min
        Q1 = $q1
        Median = $median
        Q3 = $q3
        Max = $max
        Avg = $avg
        Values = $DataArray
    }
}

$BoxSeries = New-Object Series
$BoxSeries.Name = 'BoxPlots'
$BoxSeries.ChartType = [SeriesChartType]::BoxPlot
$BoxSeries.ChartArea = $ChartArea.Name
$BoxSeries['BoxPlotShowAverage'] = 'true'
$BoxSeries['BoxPlotShowMedian'] = 'true'
$BoxSeries['BoxPlotShowInnerPoints'] = 'true'
$BoxSeries['BoxPlotShowUnusualValues'] = 'true'
$BoxSeries['BoxPlotShowExtremeValues'] = 'true'
$BoxSeries['BoxPlotShowExtremeValues'] = 'true'
$BoxSeries['BoxPlotWhiskerPercentile'] = '10'
$BoxSeries.BorderColor = 'LightGray'
$BoxSeries.BorderWidth = 2
$Chart.Series.Add($BoxSeries)

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
    # Get stats for each field
    $statistics = Get-BoxPlotStatistics -DataArray $field_values

    # Create boxplot data point
    $BoxPoint = New-Object DataPoint
    $BoxPoint.YValues = @($statistics.Min, $statistics.Q1,
        $statistics.Median, $statistics.Q3, $statistics.Max, $statistics.Avg)
    $BoxPoint.Color = $data_colors[$fields_index % $data_colors.Length]
    $BoxPoint.AxisLabel = $field

    $BoxSeries.Points.Add($BoxPoint)

    # Create series for data points
    $PointSeries = New-Object Series
    $PointSeries.Name = "Points_$field"
    $PointSeries.ChartType = [SeriesChartType]::Point
    $PointSeries.ChartArea = $ChartArea.Name
    $PointSeries.MarkerStyle = [MarkerStyle]::Circle
    $PointSeries.MarkerSize = 6
    $PointSeries.Color = 'Black'
    
    $Chart.Series.Add($PointSeries)

    foreach ($stat in $statistics.Values) {
        [void]$PointSeries.Points.AddXY($fields_index + 1, $stat)
    }

    $Label = New-Object CustomLabel
    $Label.FromPosition = $fields_index + 1 - 0.5
    $Label.ToPosition = $fields_index + 1 + 0.5
    $Label.Text = $field

    $ChartArea.AxisX.CustomLabels.Add($Label)

    $fields_index++
}

#region Boilerplate
[void]$Chart.Titles.Add('Distribution of Numeric Fields')
$Chart.Titles[0].Font = New-Object System.Drawing.Font('Lucida Console', 16)
$ChartArea.AxisX.Title = 'Numeric Fields'
$ChartArea.AxisY.Title = 'BoxPlot Statistics'
$ChartArea.AxisX.Minimum = 0

$ChartArea.AxisX.MajorGrid.LineWidth = 0
$ChartArea.AxisY.MajorGrid.LineWidth = 0

$ChartArea.AxisY.LabelStyle.ForeColor = 'Gray'

$ChartArea.AxisX.LineColor = 'LightGray'
$ChartArea.AxisX.LineWidth = 2
$ChartArea.AxisY.LineColor = 'LightGray'
$ChartArea.AxisY.LineWidth = 2

$ChartArea.AxisX.MajorTickMark.LineColor = 'LightGray'
$ChartArea.AxisY.MajorTickMark.LineColor = 'LightGray'
    
$Chart.Dock = 'Fill'  # chart adjusts to fit the entire container when the Form is resized

$Form.ShowDialog()
#endregion