<#
PS4DS: Clean & Transform Data
Author: Eric K. Miller
Last updated: 30 October 2025

This script contains PowerShell code for cleaning data.
#>

#========================
#   Cleaning functions
#========================

function Edit-Headers {
    <#
    .SYNOPSIS
        Given a data object and its source, re-import the object with
    cleaned headers.

    .DESCRIPTION
        This function gets the header names of a data object and trims
    whitespace, replaces unneeded characters, and reimports the data
    using the cleaned headers. In PowerShell, headers cannot be
    dynamically renamed, so this two-step process is required for any
    header modifications.
    
    .PARAMETER DataObject
        The data object with headers to clean.

    .PARAMETER DataSource
        The data source to reference during the reimport.
    
    .EXAMPLE
        $DataObject = Edit-Headers -DataObject $Data -DataSource "$ps4dsdata\$csv1"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$DataObject,
        [Parameter(Mandatory)]$DataSource
    )

    $cleaned_headers = ($DataObject[0].psobject.Properties).Name.Trim() -replace '\s\[.*\]','' -replace '\s+','__' -replace '\W',''
    $DataObject = Import-Csv $DataSource -Encoding UTF8 -Header $cleaned_headers
    $DataObject = $DataObject[1..$DataObject.Length]
    return $DataObject
}

function Remove-Whitespace {
    <#
    .SYNOPSIS
        Given a data object, trim all whitespace in the data.

    .DESCRIPTION
        This function loops through all headers and all rows in a data
    object to trim whitespace from the data.
    
    .PARAMETER DataObject
        The data object with data to clean.
    
    .EXAMPLE
        $DataObject = Remove-Whitespace -DataObject $Data
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$DataObject
    )

    foreach ($h in $DataObject[0].psobject.Properties.Name) {
        for ($i=0; $i -lt $DataObject.Length; $i++) {
            $DataObject[$i].$h = $DataObject[$i].$h.Trim()
        }
    }
    return $DataObject
}

function Get-InterpolatedValue {
    <#
    .SYNOPSIS
        Calculate values from a data field to use for filling in missing
    values (interpolation).

    .DESCRIPTION
        This function calculates a property of a data column and uses
    the value to set null/empty/missing values with this standardized
    value.
    
    .PARAMETER DataObject
        The data object with null values.

    .PARAMETER Field
        The field on which to interpolate values.
    
    .EXAMPLE
        $interpolatedValue = Get-InterpolatedValue -DataObject $Data -Field "height" -Measure "Average"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $DataObject,
        [Parameter(Mandatory)]
        [string]$Field,
        [Parameter(Mandatory)]
        [ValidateSet("Average", "Sum", "Maximum", "Minimum")][string]$Measure
    )

    begin {
        $filteredValues = @()
    }
    process {
        $filteredValues += $DataObject | ? {$_.$Field -notin @($null, '', 'na', 'n/a')}
        $measuredValues = $filteredValues | Measure-Object -Property $Field -Average -Sum -Maximum -Minimum
    }
    end {
        $interpolatedValue = [Math]::Floor($measuredValues.$Measure)
        return $interpolatedValue
    }
}


#========================
#  Transform functions
#========================

function ConvertTo-ColumnMultiple {
    <#
    .SYNOPSIS
        Multiply a data object column by a multiplier.
    
    .DESCRIPTION
        This function replaces values in a data object with their
    multiplied value by looping through the data and reassigning the
    values. PowerShell is not vectorized, so we need to re-calculate
    each value and reassign it to its entry to return a scaled value.

    .PARAMETER DataObject
        The data object with a column(s) to be multiplied.

    .PARAMETER Column
        The name of the column to be scaled.

    .PARAMETER Multiplier
        The factor by which to scale the column values.
    
    .EXAMPLE
        $DataObject = ConvertTo-ColumnMultiple -DataObject $Data -Column 'Numbers' -Multiplier 10
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$DataObject,
        [Parameter(Mandatory)]$Column,
        [Parameter(Mandatory)][double]$Multiplier
    )

    for ($i=0; $i -lt $DataObject.Length; $i++) {
        # Use script block for calculation
        $DataObject[$i].$Column = [regex]::Replace($DataObject[$i].$Column, '\d+(\.\d+)?',
                                    {param($match) ([Math]::Round([double]$match.Value*$Multiplier)).ToString()})
    }
    return $DataObject
}

function Get-RangeMean {
    <#
    .SYNOPSIS
        With data that has a range, calculate and replace with a single
    value from the mean.

    .DESCRIPTION
        This function checks for dashes by their unicode codes to
    identify values with ranges. It splits on this character to get min
    and max values and then calculates the mean with a specified number
    of fractional digits.

    .PARAMETER InputValue
        Data from which to calculate the mean.

    .PARAMETER FractionalDigits
        Number of significant digits to round the result (default 2).
    
    .EXAMPLE
        $RangeAvg = Get-RangeMean -InputValue $DataRow.'ColumnName'

    .NOTES
        Hat tip: https://stackoverflow.com/questions/38487021/searching-for-a-literal-hyphen-in-powershell
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][string]$InputValue,
        [Parameter()][int]$FractionalDigits=2
    )

    $Script:dash_pttrn = '[\u2010-\u2015-]'
    [double]$minrng, [double]$maxrng = $InputValue -split $Script:dash_pttrn
    #OR [double]$minrng = $InputValue -split $dash_pttrn | Select -First 1
    #   [double]$maxrng = $InputValue -split $dash_pttrn | Select -Last 1
    #OR [double]$minrng = ($InputValue -split $dash_pttrn)[0]
    #   [double]$maxrng = ($InputValue -split $dash_pttrn)[1]
    [double]$RangeAverage = [Math]::Round(($minrng+$maxrng)/2, $FractionalDigits)
    return $RangeAverage

}
