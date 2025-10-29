<#
PS4DS: Acquire (Websites)
Author: Eric K. Miller
Last updated: 28 October 2025

This script contains PowerShell code for acquiring data from websites.
#>

#========================
#   Acquire functions
#========================

function Import-Table {
    <#
    .SYNOPSIS
        Given a data table from a website, import the data into a
    PowerShell object.

    .DESCRIPTION
        This function uses a given table to extract its header and data
    and iterates through the data rows to generate a PowerShell data
    object that can be further used in PowerShell.

    .PARAMETER Table
        The original table retrieved from a website.

    .PARAMETER FormatGridView
        Switch parameter to indicate whether the data will be intended
    for Format-GridView. Ensures proper header formatting if present.
    
    .EXAMPLE
        $dataTable = Import-Table -Table $Table
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$Table,
        [Parameter()][switch]$FormatGridView
    )

    # If the user wants to view the table in GridView, format the header
    # to have double underscores for proper viewing.
    if ($FormatGridView) {
        $gvHeader = $Table.rows[0].cells | Select-Object -ExpandProperty innerText
        $Header = $gvHeader | %{$_.trim() -replace '\s+', '__' -replace '\W', ''}
    }
    else {
        $Header = $Table.rows[0].cells | Select-Object -ExpandProperty innerText
    }
    $Data = $Table.rows | Select-Object -Skip 1

    $dataTable = @()  # initialize empty array
    foreach ($row in $Data) {
        # create a progress bar!
        $progressParams = @{
            Activity = "Converting HTML table to data table"
            Status = "> Converting row $($Data.IndexOf($row)+1) of $($Data.Count) ($([Math]::Round((100*($Data.IndexOf($row)+1)/$Data.Count)))%)"
            PercentComplete = (100*($Data.IndexOf($row)+1)/$Data.Count)
        }
        Write-Progress @progressParams  # "splatting"
        
        $cells = $row.cells | Select-Object -ExpandProperty innerText | %{$_ -replace '\[.*\]', ''}  # remove brackets with citations
        $dataDict = [ordered]@{}  # initialize empty ordered hash table

        for ($i=0; $i -lt $Header.Count; $i++) {
            $dataDict[$Header[$i]] = $cells[$i]
        }
        $dataTable += [PSCustomObject]$dataDict  # convert ordered hash table to PSCustomObject
    }
    return $dataTable
}

# Utility function for below
function Save-CsvFile ($InitialDirectory) {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = $InitialDirectory
    $SaveFileDialog.Filter = "CSV (*.csv) | *.csv"
    $SaveFileDialog.ShowDialog() | Out-Null
    return $SaveFileDialog.FileName
}

function Convert-WebTable {
    <#
    .SYNOPSIS
        View data from a website table in a PowerShell table, in
    GridView, or download to a CSV.

    .DESCRIPTION
        This function uses a given table to format and view the data
    in native PowerShell formats, or optionally downloads the data to
    a CSV.

    .PARAMETER Table
        The original table retrieved from a website.
    
    .EXAMPLE
        Convert-WebTable -Table $Table
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$Table
    )

    $options = ''
    while ($options.ToLower() -notin @('1','table','2','grid','3','csv')) {
        $options = Read-Host @"
Select an option for using the data:
 1 - View in table`t[Type '1' or 'table']
 2 - View in grid`t[Type '2' or 'grid']
 3 - Download CSV`t[Type '3' or 'csv']`n
"@

        switch ($options) {
            {$_ -in @('1','table')}
                {
                    $dataTable = Import-Table -Table $Table
                    $dataTable | Format-Table -AutoSize -Wrap -RepeatHeader -Expand Both; break
                }
            {$_ -in @('2','grid')}
                {
                    $dataTable = Import-Table -Table $Table -FormatGridView
                    $tableCaption = $table.caption.innerText
                    $gvTitle = "Wikipedia Data Table: $tableCaption"
                    $dataTable | Out-GridView -Title $gvTitle; break
                }
            {$_ -in @('3','csv')}
                {
                    $dataTable = Import-Table $Table
                    $CsvFilePath = Save-CsvFile('.')
                    $dataTable | Export-Csv -Path $CsvFilePath -Encoding UTF8 -NoTypeInformation
                    Write-Host "`nData written to $CsvFilePath"; break
                }
            default
                {Write-Host "`nInvalid input. Please select one of the below.`n"}
        }
    }
}