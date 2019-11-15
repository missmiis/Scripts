PARAM ($PortalConfigFolder)

<#
    ConvertToCSV-PortalConfig.ps1

    Written by Carol Wapshere, http://www.wapshere.com.missmiis

    Takes the policy.xml and schema.xml files exported and converts them to a CSV file per object type.
    
    NOTES:
    - Export policy.xml and schema.xml using the MS process at https://www.microsoft.com/en-us/download/details.aspx?id=1663
    - Produces semi-colon separated CSV files for each object type. 
    - Only attributes suitable for inclusion in the CSV are included - ie, excludes embedded XML and values with semicolons in them.
    - Multi-valued attributes are comma-seperated within the semi-colon delimued format.
    - Where an attribuite contains a GUID that is listed in the same XML file, the script will swap it for the target object's Display Name.
#>

$ErrorActionPreference = "Stop"

if (-not $PortalConfigFolder)
{
    throw "The folder where policy.xml and schema.xml are located must be specified as PortalConfigFolder"
}

$ScriptFldr = split-path -parent $MyInvocation.MyCommand.Definition

$ConfigFiles = @("schema.xml","policy.xml")

$XSLTFile = $ScriptFldr + "\MIMServiceToCSV.xslt"
$xslt = new-object system.xml.xsl.XslTransform
$xslt.load($XSLTFile)

foreach ($ConfigFile in $ConfigFiles)
{
    $SourceFile = $PortalConfigFolder + "\" + $ConfigFile
    $TempFile = $SourceFile.Replace(".xml",".csv")

    $xslt.Transform($SourceFile, $TempFile) 

    $csv = Import-csv $TempFile -Delimiter ";"
    $csv = $csv | where {$_.id  -like "urn:uuid:*"}

    $DisplayNames = @{}
    foreach ($item in ($csv | where {$_.attribute -eq "DisplayName"} | select -Property id,value)) {$DisplayNames.Add($item.id,$item.value)}

    $ObjectTypes = $csv.type | select -unique
    foreach ($ObjectType in $ObjectTypes)
    {
        $ObjectType
        $rows = $csv | where {$_.type -eq $ObjectType}
        $attributes = $rows.attribute | select -Unique
        $objectIDs = $rows.id | select -Unique
        $arr = @()
        foreach ($objectID in $objectIDs)
        {
            $objrows = $rows | where {$_.id -eq $objectID}
            $pso = New-Object PSObject
            foreach ($attribute in $attributes)
            {
                $value = $null
                try{$value = ($objrows | where {$_.attribute -eq $attribute}).value} catch {}
                if ($attribute -ne "ObjectID" -and $value -and $value.StartsWith("urn:uuid:"))
                {
                    $tmp = $value
                    foreach ($guid in $tmp.split(","))
                    {
                        if ($DisplayNames.ContainsKey($guid))
                        {
                            $value = $value.Replace($guid,($DisplayNames.($guid)))
                        }
                    }
                }

                $pso | Add-Member -MemberType NoteProperty -Name $attribute -Value $value
            }
            $arr += $pso
        }
        $arr | Export-Csv -Path ($SourceFile.Replace(".xml","") + "_" + $ObjectType + ".csv") -NoTypeInformation

    }

    Remove-Item $TempFile
}
