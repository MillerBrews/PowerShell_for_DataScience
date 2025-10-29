<#
PS4DS: Breakout SQLite
Author: Eric K. Miller
Last updated: 29 October 2025

This script contains PowerShell code for querying and updating a
SQLite database.
#>

#========================
# Functions for SQL work
#========================

using namespace System.Data

function Select-SQLiteObject {
    <#
    .SYNOPSIS
        Enables SQLite SELECT statements to pull data from a database
    into a data object.

    .DESCRIPTION
        This function enables the user to use SQL statements to query
    a specified database. The user can write the SQL as a string, or
    import a SQL statement to use for the query.
    
    .PARAMETER Database
        The database name to query.

    .PARAMETER Sql
        SQL code to execute (SELECT statement).
    
    .EXAMPLE
        $DataOutput = Select-SQLiteObject -Database "YourDBName.db" -Sql "SELECT * FROM Star_Wars ORDER BY name LIMIT 5"
    #>
    param (
        [Parameter(Mandatory)][string]$Database,
        [Parameter(Mandatory)][string]$Sql
    )

    # Load SQLite assembly if not already loaded
    if (-not ("System.Data.SQLite.SQLiteConnection" -as [type])) {
        Add-Type -Path $sqlite_dll  # $profile variable
    }

    $conn = New-Object SQLite.SQLiteConnection
    $conn.ConnectionString = "Data Source=$Database"
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Sql

    $SQLiteAdapter = New-Object SQLite.SQLiteDataAdapter
    $SQLiteAdapter.SelectCommand = $cmd
    [void]$SQLiteAdapter.SelectCommand.ExecuteNonQuery()

    $DataSet = New-Object DataSet
    $SQLiteAdapter.Fill($DataSet) | Out-Null

    <#
    # Using the ExecuteReader() method, the data is returned
    # row-by-row, instead of in an object as above.
    $reader = $cmd.ExecuteReader()

    $rows = @()
    # Read each row and build a hashtable
    while ($reader.Read()) {
        $row = [ordered]@{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $columnName = $reader.GetName($i)
            $value = $reader.GetValue($i)
            $row[$columnName] = $value
        }
        $rows += [PSCustomObject]$row
    }

    #$reader.Close()
    #>

    $conn.Close()

    return $DataSet.Tables[0]
}

function Push-SQLiteObject {
    <#
    .SYNOPSIS
        Pushes a data object to a SQLite database, creating a new table
    if it does not exist.

    .DESCRIPTION
        This function creates a new table in a specified database from
    a data object. The function uses transactions to enable bulk
    commits for performant SQL actions.
    
    .PARAMETER Database
        The database name for the table.

    .PARAMETER TableName
        User-supplied name of the table to create/update.

    .PARAMETER DataObject
        The data (a PowerShell object) to insert into a table in the
    database.
    
    .EXAMPLE
        Push-SQLiteObject -Database "YourDBName.db" -TableName "Star_Wars" -DataObject $StarWars

    .NOTES
        The function infers table field types from data object field
    types, with the default being TEXT. Therefore, it is recommended to
    ensure the data object has strongly typed data fields before
    pushing to SQLite.
    #>
    param (
        [Parameter(Mandatory)][string]$Database,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][object[]]$DataObject
    )

    # Load SQLite assembly if not already loaded
    if (-not ("System.Data.SQLite.SQLiteConnection" -as [type])) {
        Add-Type -Path $sqlite_dll  # $profile variable
    }

    # Open connection
    $conn = New-Object System.Data.SQLite.SQLiteConnection
    $conn.ConnectionString = "Data Source=$Database"
    $conn.Open()

    # Build column and parameter lists
    $columns = $DataObject[0].PSObject.Properties | % {
        $dtype = switch ($_.Value.GetType().Name) {
            "Int32" {"INTEGER"}
            "Int64" {"INTEGER"}
            "Double" {"REAL"}
            "Decimal" {"REAL"}
            "Boolean" {"INTEGER"}
            default {"TEXT"}
        }
        "$($_.Name) $dtype"
    }
    $columnList = $columns -join ", "

    # Create table if needed
    $createSQL = "CREATE TABLE IF NOT EXISTS $TableName ($columnList);"
    $createCmd = $conn.CreateCommand()
    $createCmd.CommandText = $createSQL
    [void]$createCmd.ExecuteNonQuery()  # suppress output

    # Prepare insert
    $colNames = $DataObject[0].PSObject.Properties.Name
    $paramNames = $colNames | %{"@$_"}
    $insertSQL = "INSERT INTO $TableName (" + ($colNames -join ", ") + ") VALUES (" + ($paramNames -join ", ") + ");"
    
    # Begin transaction
    $transaction = $conn.BeginTransaction()

    foreach ($row in $DataObject) {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $insertSQL
        $cmd.Transaction = $transaction

        foreach ($col in $colNames) {
            $value = $row.$col
            $cmd.Parameters.AddWithValue("@$col", $value) | Out-Null  # suppress output
        }
        [void]$cmd.ExecuteNonQuery()
    }
    # Commit transaction
    $transaction.Commit()
    $conn.Close()
}
