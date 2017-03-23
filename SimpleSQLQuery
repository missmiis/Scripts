Function Invoke-SQLQuery
{
<#
.SYNPOSIS
    Connects to a SQL Database using integrated security and runs a SQL Query
.OUTPUT
    An array of PSObjects, one for each row returned by the query
.PARAMETERS
    Server: The name of the SQL server
    Instance: The SQL Server Instance
    Database: The name of the Database
    ConnectionTimeout: Timeout in seconds
    Query: The query to run
.DEPENDENCIES
    The account running the script must have access to the DB.
.ChangeLog
    Carol Wapshere, 23/2/2017, Adapted from http://powershell4sql.codeplex.com
#>

   PARAM (
    [Parameter(Mandatory=$true)][string] $Server, 
    [Parameter(Mandatory=$true)][string] $Database, 
    [string] $Instance,
    [int] $ConnectionTimeout = 60,
            [Parameter(Mandatory=$true)] [String]$Query
    ) 
    END
    {
        ## Construct connection string
        if ($Instance -and $Instance -ne "" -and $Instance -ne "DEFAULT")
        {
            $SQLServer = $Server + "\" + $Instance
        }
        else
        {
            $SQLServer = $Server
        }
        $cs = "Data Source=$SQLServer;Initial Catalog=$Database;Connect Timeout=$ConnectionTimeout;Integrated Security=true;"
        Write-Verbose -Message "Connection String: $cs"

        ## Open SQL connection
        Try
        {
            $SqlConnection = new-object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = $cs
            $SqlConnection.Open()
            Write-Verbose -Message  "SQL Connection opened"
        }
        Catch
        {
            Write-Error -Message ("Failed to connect to SQL Database $Database. " + $_.Exception.Message)
        }

        ## Run the query
        Try
        {
                $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
                $SqlCmd.CommandText = $Query
                $SqlCmd.Connection = $SqlConnection
                $reader = $SqlCmd.ExecuteReader()
            Write-Verbose -Message "Query complete, reading data"
        }
        Catch
        {
            Write-Error -Message ("Failed to run SQL query. " + $_.Exception.Message)
        }


        ## Read the query response into a data array
                $data = @()
        if ($reader.HasRows)
        {
            do {
                $recordsetIndex++;
                [int] $rec = 0;

                while ($reader.Read()) {
                    $rec++;
                    $record = New-object PSObject

                    for ($i = 0; $i -lt $reader.FieldCount; $i++)
                    {
                        $name = $reader.GetName($i);
                        if ([string]::IsNullOrEmpty($name))
                        {
                            $name = "Column#" + $i.ToString();
                        }
                            
                        $val = $reader.GetValue($i);
                            
                        Add-Member -MemberType NoteProperty -InputObject $record -Name $name -Value $val 
                    }
                        
                    if ($IncludeRecordSetIndex)
                    {
                        Add-Member -MemberType NoteProperty -InputObject $record -Name "RecordSetIndex" -Value $recordsetIndex
                    }

                    $data += $record
                }                                                   
                           
            } while ($reader.NextResult())

            $reader.Close()

            if ($data) 
            {
                Write-Verbose -Message ("Returning {0} rows" -f $data.count)
            }
            else
            {
                Write-Error -Message "Failed to parse returned data"
            }
        }
        else
        {
            Write-Verbose -Message "Query returned 0 rows"
        }

        ## Close the SQL connection
        $SqlConnection.Close()

        Write-Verbose "Done."
        Return $data
    }
}
