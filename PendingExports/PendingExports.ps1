PARAM($MAName,$ObjectType,$CSExportFile,$ReportFolder,$RemoveLineBreaksFromCSExport=$true)

## 
## PendingExports.ps1
## Written by Carol Wapshere, www.wapshere.com/missmiis
##
## Exports the pending exports from the connector space to an XML file then converts to two CSV files - one for single-valued and one for multi-valued attributes.
## First splits the XML into multiple smaller files (one for each cs-object) to allow processing of large csexport XML files.
##
## USAGE:  .\PendingExports.ps1 -MAName "MA Name" -ObjectType "type" [-CSExportFile "full path"] [-ReportFolder "folder to save CSV files"]
##
## NOTES:
##   - Needs CSExportSplitLines.xslt in the same folder as this script
##   - Either run on the MIM Sync server or use csexport with the "/f:x" filter to export the pending exports and pass the file name.
##   - If no ReportFolder specified the generated files are saved to the folder this script was run from.
##   - If any attributes have carriage returns in the value it messes up the CSV - set $RemoveLineBreaksFromCSExport to $true
##   - This script is designed to be run after a Full Sync and before running any exports. It will fail if there are unconfirmed exports.
##   - Update $ExcludeAttribs to list any attributes that should be excluded from the reports.
##


$ErrorActionPreference = "Stop"

$CSExportExe = "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Bin\csexport.exe"

$ScriptFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
if (-not $ReportFolder) {$ReportFolder = $ScriptFolder}

$ExcludeAttribs = @("objectSid","whenCreated","pwdLastSet","lastLogonTimestamp","homeMDB")

#region Export Connector Space
if (-not $CSExportFile)
{
    $CSExportFile = $ReportFolder + "\" + $MAName + "_CS.xml"
    if (Test-Path $CSExportFile) {Remove-Item $CSExportFile}
    . $CSExportExe $MAName $CSExportFile /f:x
}
#endregion Export Connector Space


#region Break up XML into CSObjects
if ($RemoveLineBreaksFromCSExport)
{
    $Reader = [System.IO.File]::ReadAllText($CSExportFile)
    $CSExportFile = $CSExportFile.Replace(".xml","1.xml")
    $Reader.Replace("`n","").Replace("`r","") | out-file $CSExportFile
}

## Break up the CSExport file
write-host "`nCreating temporary working folder..."
$TempFolder = $ReportFolder + "\" + $MAName + "_" + $ObjectType + ".TMP"
if (Test-Path $TempFolder)
{
    Remove-Item $TempFolder -Force -Recurse
}
New-Item $TempFolder -ItemType Directory

## The CSExport file has no line breaks - the first step inserts them so each cs-object is a seperate line.
write-host "`nCopying the CSExport file to a temporary version with line breaks between each cs-object node..."
$XSLTPath = $ScriptFolder + "\CSExportSplitLines.xslt"
$SourceFile = $CSExportFile
$TargetFile = $TempFolder + "\" + $MAName + ".XML"

$xslt = new-object system.xml.xsl.XslTransform
$xslt.load($XSLTPath)
$xslt.Transform($SourceFile,$TargetFile)   

## Read the new XML file line by line and split out into cs-object files.
write-host "`nSplitting the CSExport file into per-object temporary files..."
$reader = [System.IO.File]::OpenText($TargetFile)
$i = 0
$ObjectTypeCheck = "object-type=""" + $ObjectType + """"
if ($reader -and -not $reader.EndOfStream) {do {
    $line = $reader.ReadLine()
    if ($line -and $line.Contains($ObjectTypeCheck))
    {
        $line | out-file ($TempFolder + "\" + $i.ToString().PadLeft(10,'0') + ".XML")
        $i += 1
    }
} until ($reader.EndOfStream)}

$reader.Close()
Remove-Item $TargetFile

write-host "`nFound $i $ObjectType objects with pending changes. `n"

#endregion Break up XML into CSObjects


#region Parse each cs-object file

write-host "Parsing each cs object..."

$CSObjects = @{}
$SingleAttribs = @{}
$MultiAttribs = @{}

foreach ($CSObjectFile in (Get-ChildItem $TempFolder -Filter *.XML))
{
    $CSObjectFile.FullName
    [xml]$CSData = Get-Content $CSObjectFile.FullName
    $CSObj = $CSData."cs-object"

    $DN = $CSObj."cs-dn"
    $DN
    $CSObjects.Add($DN,@{})

    if ($CSObj."unapplied-export".delta.operation)
    {
        $CSObjects.($DN).Add("operation",$CSObj."unapplied-export".delta.operation)
    }
    else
    {
        $CSObjects.($DN).Add("operation","none")
    }

    ##
    ## DN Change
    ##
    if ($CSObj."unapplied-export".delta.newdn)
    {
        $CSObjects.($DN).Add("newdn",$CSObj."unapplied-export".delta.newdn)
    }

#region Existing values
    foreach ($attr in $CSObj."synchronized-hologram".entry.attr)
    {
        if (-not $CSObjects.($DN).ContainsKey($attr.name)) {$CSObjects.($DN).Add($attr.name,@{})}
        if ($attr.multivalued -eq "false")
        {
            if (-not $SingleAttribs.ContainsKey($attr.name)) {$SingleAttribs.Add($attr.name,0)}
            $CSObjects.($DN).($attr.name).Add("current",$attr.value)
        }
        else
        {
            if (-not $MultiAttribs.ContainsKey($attr.name)) {$MultiAttribs.Add($attr.name,0)}
            $CSObjects.($DN).($attr.name).Add("current",($attr.value -join ";"))
        }
    }
    foreach ($attr in $CSObj."synchronized-hologram".entry."dn-attr")
    {
        if (-not $CSObjects.($DN).ContainsKey($attr.name)) {$CSObjects.($DN).Add($attr.name,@{})}
        if ($attr.multivalued -eq "false")
        {
            if (-not $SingleAttribs.ContainsKey($attr.name)) {$SingleAttribs.Add($attr.name,0)}
            $CSObjects.($DN).($attr.name).Add("current",$attr."dn-value".dn)
        }
        elseif ($attr.name)
        {
            if (-not $MultiAttribs.ContainsKey($attr.name)) {$MultiAttribs.Add($attr.name,0)}
            $CSObjects.($DN).($attr.name).Add("current",($attr."dn-value".dn))
        }        
    }
#endregion Existing values


#region Attribute Changes
    if ($CSObj."unapplied-export".delta.attr)
    {
        foreach ($attr in $CSObj."unapplied-export".delta.attr)
        {
            if (-not $CSObjects.($DN).ContainsKey($attr.name)) {$CSObjects.($DN).Add($attr.name,@{})}
            if ($attr.multivalued -eq "false")
            {
                if (-not $SingleAttribs.ContainsKey($attr.name)) {$SingleAttribs.Add($attr.name,1)} else {$SingleAttribs.($attr.name) += 1}

                if ($attr.operation -eq "delete") {$CSObjects.($DN).($attr.name).Add("delete",$null)}
                elseif ($attr.operation -eq "add") {$CSObjects.($DN).($attr.name).Add("add",$attr.value)}
                elseif ($attr.value.count -gt 1) {$CSObjects.($DN).($attr.name).Add("add",($attr.value | where {$_.operation -eq "add"})."#text")}
                elseif ($attr.value) {$CSObjects.($DN).($attr.name).Add("add",$attr.value)}
            } 
            elseif ($attr.multivalued -eq "true")
            {
                if (-not $MultiAttribs.ContainsKey($attr.name)) {$MultiAttribs.Add($attr.name,1)} else {$MultiAttribs.($attr.name) += 1}

                if ($attr.operation -eq "delete") {$CSObjects.($DN).($attr.name).Add("delete",$null)}
                elseif ($attr.operation -eq "add") {$CSObjects.($DN).($attr.name).Add("add",$attr.value)}
                elseif ($attr.value.Count -gt 1) 
                {
                    $CSObjects.($DN).($attr.name).Add("add",($attr.value | where {$_.operation -eq "add"})."#text" -join ";")
                    $CSObjects.($DN).($attr.name).Add("delete",($attr.value | where {$_.operation -eq "delete"})."#text" -join ";")
                }
                elseif ($attr.value) {$CSObjects.($DN).($attr.name).Add("add",$attr.value)}
            } 

        }
    }
#endregion Attribute Changes


#region Reference Attribute Changes
    if ($CSObj."unapplied-export".delta."dn-attr")
    {
        foreach ($attr in $CSObj."unapplied-export".delta."dn-attr")
        {
            if (-not $CSObjects.($DN).ContainsKey($attr.name)) {$CSObjects.($DN).Add($attr.name,@{})}
            if ($attr.multivalued -eq "false")
            {
                if (-not $SingleAttribs.ContainsKey($attr.name)) {$SingleAttribs.Add($attr.name,1)} else {$SingleAttribs.($attr.name) += 1}

                if ($attr.operation -eq "delete") {$CSObjects.($DN).($attr.name).Add("delete",$null)}
                elseif ($attr.operation -eq "add") {$CSObjects.($DN).($attr.name).Add("add",$attr."dn-value".dn)}
                elseif ($attr."dn-value") {$CSObjects.($DN).($attr.name).Add("add",($attr.value | where {$_.operation -eq "add"})."dn-value".dn)}
            } 
            elseif ($attr.multivalued -eq "true")
            {
                if (-not $MultiAttribs.ContainsKey($attr.name)) {$MultiAttribs.Add($attr.name,1)} else {$MultiAttribs.($attr.name) += 1}

                if ($attr."dn-value")
                {
                    $CSObjects.($DN).($attr.name).Add("add",($attr."dn-value" | where {$_.operation -eq "add"}).dn)
                    $CSObjects.($DN).($attr.name).Add("delete",($attr."dn-value" | where {$_.operation -eq "delete"}).dn)
                }
            } 
        }
    }
#endregion Reference Attribute Changes

}


Remove-Item $TempFolder -Force -Recurse
write-host "`n"

#endregion Parse each cs-object file


#region Output Report files

## Single values
$CSVSingle = $ReportFolder + "\" + $MAName + "_" + $ObjectType + "_SingleValue.csv"
$Header = "DN;operation;DN-new"
$Content = @()

foreach ($attrib in $SingleAttribs.Keys | sort) {if ($ExcludeAttribs -notcontains $attrib){$Header = $Header + ";$attrib-current;$attrib-change"}}
$Header | Out-File $CSVSingle -Encoding Default

$j = 0
foreach ($DN in $CSObjects.Keys)
{
    $CSVLine = $DN + ";" + $CSObjects.($DN).operation
    if ($CSObjects.($DN).ContainsKey("newdn")){$CSVLine = $CSVLine + ";" + $CSObjects.($DN).newdn}
    else {$CSVLine = $CSVLine + ";"}

    foreach ($attrib in $SingleAttribs.Keys | sort)
    {
        if ($ExcludeAttribs -notcontains $attrib)
        {
            $CSVLine = $CSVLine + ";" + $CSObjects.($DN).($attrib).current + ";"
            if ($CSObjects.($DN).ContainsKey($attrib))
            {
                if ($CSObjects.($DN).($attrib).ContainsKey("add")) {$CSVLine = $CSVLine + $CSObjects.($DN).($attrib)."add"}
                elseif ($CSObjects.($DN).($attrib).ContainsKey("delete")) {$CSVLine = $CSVLine + "**DELETE**"}
            }
        }
    }
    #$CSVLine
    $CSVLine | Add-Content $CSVSingle

    $j += 1
    write-progress -id  1 -activity Updating -status 'Writing single-value CSV file' -percentcomplete ($j/$i * 100)
}

write-host "Single-valued pending exports saved to: $CSVSingle"


## Multi values
$CSVMulti = $ReportFolder + "\" + $MAName + "_" + $ObjectType + "_MultiValue.csv"
$Header = "DN;attribute;operation;value"
$Content = @()

$Header | Out-File $CSVMulti -Encoding Default

$j = 0
foreach ($DN in $CSObjects.Keys)
{
    if ($MultiAttribs.Count -gt 0) 
    {
        foreach ($attrib in $MultiAttribs.Keys)
        {
           foreach ($change in $CSObjects.($DN).($attrib).Keys)
           {
                foreach ($value in $CSObjects.($DN).($attrib).($change))
                {
                    $CSVLine = $DN + ";" + $attrib + ";" + $change + ";" + $value
                    $CSVLine | Add-Content $CSVMulti
                }
            }
        }
        #$CSVLine
        #$CSVLine | Add-Content $CSVMulti
    }

    $j += 1
    write-progress -id  2 -activity Updating -status 'Writing multi-value CSV file' -percentcomplete ($j/$i * 100)
}

write-host "Multi-valued pending exports saved to: $CSVMulti"



## Summary Report on group membership changes 
if ($ObjectType -eq "group")
{
    $CSVGroup = $ReportFolder + "\" + $MAName + "_" + $ObjectType + "_GroupMemberChanges.csv"
    $Header = "DN;Current;Adds;Deletes"
    $Content = @()

    write-host "Writing summary CSV of group membership changes..."
   $Header | Out-File $CSVGroup -Encoding Default

    $j = 0
    foreach ($DN in $CSObjects.Keys)
    {
        if ($CSObjects.($DN).member.current.count -gt 0 -or $CSObjects.($DN).member.add.count -gt 0 -or $CSObjects.($DN).member.delete.count -gt 0) 
        {
            $CSVLine = $DN + ";" + $CSObjects.($DN).member.current.count.ToString() + ";" + $CSObjects.($DN).member.add.count.ToString() + ";" + $CSObjects.($DN).member.delete.count.ToString()
            $CSVLine | Add-Content $CSVGroup
        }

        $j += 1
        write-progress -id  3 -activity Updating -status 'Writing membership change CSV file' -percentcomplete ($j/$i * 100)
    }
    write-host "Group membership changes summary report saved to: $CSVGroup"

}


##
## Write HTML Summary Report
##

$ReportFile = $ReportFolder + "\" + $MAName + "_" + $ObjectType + "_Summary.html"

"<html><body>" | Out-File $ReportFile -Encoding Default
"<h2>" + $MAName + " " + $ObjectType + " Pending Exports at " + (get-item $CSExportFile).CreationTime + "</h2>" | Add-Content $ReportFile
"<p>Total Pending Exports: " + $CSObjects.Count.ToString() | Add-Content $ReportFile

"<p>Number of objects with changes per attribute: " | Add-Content $ReportFile
"<ul>" | Add-Content $ReportFile

foreach ($attrib in $SingleAttribs.Keys | sort) {if ($SingleAttribs.($attrib) -gt 0)
{
    "<li>" + $attrib + ": " + $SingleAttribs.($attrib).ToString() + "</li>" | Add-Content $ReportFile
}}
foreach ($Attrib in $MultiAttribs.Keys | sort)
{
    "<li>" + $Attrib + ": " + $MultiAttribs.($Attrib).ToString() + "</li>" | Add-Content $ReportFile
}
"</ul>" | Add-Content $ReportFile

<#
"<p>For a full list of pending export changes see:" | Add-Content $ReportFile
"<ul>" | Add-Content $ReportFile
if ($SingleAttribs.Count -gt 0) {"<li>" + $MAName + "_" + $ObjectType + "_SingleValue.csv" + "</li>" | Add-Content $ReportFile}
if ($MultiAttribs.Count -gt 0) {"<li>" + $MAName + "_" + $ObjectType + "_MultiValue.csv" + "</li>" | Add-Content $ReportFile}
if ($GroupChangeSummaryReport -and $ObjectType -eq "group") {"<li>" + $MAName + "_" + $ObjectType + "_GroupMemberChanges.csv" + "</li>" | Add-Content $ReportFile}
"</ul>" | Add-Content $ReportFile
#>
"</body></html>" | Add-Content $ReportFile

write-host "HTML summary report saved to: $ReportFile"

#endregion Output Report files

