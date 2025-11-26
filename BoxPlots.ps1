using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

# Reference: https://learn.microsoft.com/en-us/previous-versions/dd456709(v=vs.140)

#region Data setup
<#
$StarWars = Import-Csv '.\Documents\PowerShell_for_Data_Science\Data\StarWars.csv' -Encoding UTF8

# Ensure numeric columns are strongly typed
$numeric_headers = @('height', 'mass')
foreach ($h in $numeric_headers) {
    for ($i=0; $i -lt $StarWars.Length; $i++) {
        if ($StarWars[$i].$h -ne '') {
            $StarWars[$i].$h = [double]$StarWars[$i].$h
        }
        else {$StarWars[$i].$h = $null}
    }
}
#>
# Get numeric fields
$numeric_fields = $StarWars[0].PSObject.Properties |
    Where-Object {$_.Value -is [double] -or $_.Value -is [decimal]} |
    ForEach-Object {$_.Name}

# Select data to use
$StarWars_NoJabba = $StarWars | ?{$_.mass -lt 200 -and $_.mass -ne $null}
#endregion

#region Boilerplate charting objects creation
# Chart to hold the ChartArea and Series
$Chart = New-Object Chart
$Chart.Width = 700
$Chart.Height = 500
$ChartArea = New-Object ChartArea
$ChartArea.Name = 'ChartArea'
$Chart.ChartAreas.Add($ChartArea)

# Form to hold the chart
$Form = New-Object Form
$Form.Width = 700
$Form.Height = 500
$Form.Controls.Add($Chart)
$Form.Text = 'BoxPlot'
#endregion

<# MANUAL BOXPLOTS
$BoxSeries = New-Object Series
$BoxSeries.ChartType = [SeriesChartType]::BoxPlot
$BoxSeries.ChartArea = $ChartArea.Name
$BoxSeries['PointWidth'] = '0.5'
$BoxSeries.Palette = 'BrightPastel'
#None, Bright, Grayscale, Excel, Light, Pastel, EarthTones,
#SemiTransparent, Berry, Chocolate, Fire, SeaGreen, BrightPastel
$BoxSeries.BorderColor = 'Black'
$BoxSeries.BorderWidth = 1
$BoxSeries.ToolTip = "BoxPlot Statistics:\n\n" +
                     "Min Whisker: #VALY1{0.0}\n" +
                     "Q1 (25%): #VALY3{0.0}\n" +
                     "Avg: #VALY5{0.0}\n" +
                     "Median (- - -): #VALY6{0.0}\n" +
                     "Q3 (75%): #VALY4{0.0}\n" +
                     "Max Whisker: #VALY2{0.0}\n"
#>

$fields_index = 0
$PointSeriesNames = @()
foreach ($field in $numeric_fields) {
    # Extract values for each field
    $field_values = $StarWars_NoJabba.$field
    
    # Point data series
    $PointSeries = New-Object Series
    $PointSeries.Name = "PointData_$fields_index"
    $PointSeries.ChartType = [SeriesChartType]::Point
    $PointSeries.ChartArea = $ChartArea.Name
    $PointSeries.MarkerStyle = [MarkerStyle]::Circle
    $PointSeries.MarkerSize = 6
    $PointSeries.Color = 'Gray'
    [double[]]$field_values | %{[void]$PointSeries.Points.AddXY($fields_index + 1, $_)}
    $PointSeries.ToolTip = "$($field): #VALY1{0.0}"
    $Chart.Series.Add($PointSeries)

    $PointSeriesNames += $PointSeries.Name  # collect for automatic boxplots
    
    <# BOXPOINTS FOR MANUAL BOXPLOTS
    $field_values_sorted = $field_values | Sort-Object
    $data_length = $field_values_sorted.Count

    # Six stats values calculations
    $median = if ($data_length % 2 -eq 0) {
            ($field_values_sorted[$data_length / 2] +
            $field_values_sorted[$data_length / 2 - 1]) / 2
        }
        else { $field_values_sorted[[Math]::Floor($data_length / 2)] }
    $Q1 = $field_values_sorted[[Math]::Floor($data_length * 0.25)]
    $Q3 = $field_values_sorted[[Math]::Floor($data_length * 0.75)]
    $IQR = $Q3-$Q1
    $min_whisker = $field_values_sorted[[Math]::Floor($data_length * 0.10)]  # $Q1 - 1.5*$IQR  #
    $max_whisker = $field_values_sorted[[Math]::Floor($data_length * 0.90)]  # $Q3 + 1.5*$IQR  #
    $avg = ($field_values_sorted | Measure-Object -Average).Average

    # Create BoxPlot data points
    $BoxPoints = New-Object DataPoint
    $stats = @($min_whisker, $max_whisker, $Q1, $Q3, $avg, $median)
    $BoxPoints.SetValueXY($fields_index + 1, $stats)
    $BoxSeries.Points.Add($BoxPoints)
    #>
    
    $Label = New-Object CustomLabel
    $Label.FromPosition = $fields_index + 1 - 0.5
    $Label.ToPosition = $fields_index + 1 + 0.5
    $Label.Text = $field
    $ChartArea.AxisX.CustomLabels.Add($Label)

    $fields_index++
}
# AUTOMATIC BOXPLOTS
$BoxSeries = New-Object Series
$BoxSeries.ChartType = [SeriesChartType]::BoxPlot
$BoxSeries.ChartArea = $ChartArea.Name
# BoxPlot point values are calculated and added for each series,
# delimited by semicolons
$BoxSeries['BoxPlotSeries'] = $PointSeriesNames -join ';'
$BoxSeries['BoxPlotPercentile'] = '25'  # default
$BoxSeries['BoxPlotWhiskerPercentile'] = '10'  # default
$BoxSeries['BoxPlotShowUnusualValues'] = 'True'  # default
$BoxSeries['PointWidth'] = '0.5'
$BoxSeries.Palette = 'BrightPastel'
$BoxSeries.BorderColor = 'Black'
$BoxSeries.BorderWidth = 1
$BoxSeries.ToolTip = "BoxPlot Statistics:\n\n" +
                     "Min Whisker: #VALY1{0.0}\n" +
                     "Q1 (25%): #VALY3{0.0}\n" +
                     "Avg: #VALY5{0.0}\n" +
                     "Median (- - -): #VALY6{0.0}\n" +
                     "Q3 (75%): #VALY4{0.0}\n" +
                     "Max Whisker: #VALY2{0.0}\n"
#>
$Chart.Series.Add($BoxSeries)

#region Chart settings
# Titles
[void]$Chart.Titles.Add('Distribution of Numeric Fields')
$Chart.Titles[0].Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 16)
$ChartArea.AxisX.Title = 'Numeric Fields'
$ChartArea.AxisX.TitleFont = New-Object System.Drawing.Font('Microsoft Sans Serif', 11)
$ChartArea.AxisX.TitleForeColor = 'Gray'
$ChartArea.AxisY.Title = 'BoxPlot Statistics'
$ChartArea.AxisY.TitleFont = New-Object System.Drawing.Font('Microsoft Sans Serif', 11)
$ChartArea.AxisY.TitleForeColor = 'Gray'

# AxisX
$ChartArea.AxisX.LineColor = 'LightGray'
$ChartArea.AxisX.LineWidth = 2
$ChartArea.AxisX.Minimum = 0
$ChartArea.AxisX.Interval = 1
$ChartArea.AxisX.MajorTickMark.LineColor = 'LightGray'
$ChartArea.AxisX.MajorGrid.LineWidth = 0

# AxisY
$ChartArea.AxisY.LineColor = 'LightGray'
$ChartArea.AxisY.LineWidth = 2
$ChartArea.AxisY.MajorTickMark.LineColor = 'LightGray'
$ChartArea.AxisY.MajorGrid.LineWidth = 0
$ChartArea.AxisY.LabelStyle.ForeColor = 'Gray'

$ChartArea.BackColor = 'LightGray'

# Chart adjusts to fit the entire container when the Form is resized
$Chart.Dock = 'Fill'

$Form.ShowDialog()
#endregion