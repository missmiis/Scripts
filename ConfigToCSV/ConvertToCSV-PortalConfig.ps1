PARAM ($PortalConfigFolder)

<#
    ConvertToCSV-PortalConfig.ps1

    Written by Carol Wapshere, http://www.wapshere.com.missmiis

    Takes the policy.xml and schema.xml files exported using the standard MIM Service migration scripts, see https://www.microsoft.com/en-us/download/details.aspx?id=1663
    Produces semi-colon separated CSV files for each object type. Only attributes suitable for inclusion in the CSV are included (ie excluded embedded XML and values with semicolons in).
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
                $pso | Add-Member -MemberType NoteProperty -Name $attribute -Value ($objrows | where {$_.attribute -eq $attribute}).value
            }
            $arr += $pso
        }
        $arr | Export-Csv -Path ($SourceFile.Replace(".xml","") + "_" + $ObjectType + ".csv") -NoTypeInformation

    }

    Remove-Item $TempFile
}
