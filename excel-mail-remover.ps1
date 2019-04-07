param (
    [string]$unsubscribedSuffix = '_unsuscribed_members',
    [string]$subscribedSuffix = '_subscribed_members',
    [string]$directory = $PSScriptRoot,
    [switch]$update = $false,
    [string]$reportXls = ''
 )

 function GetEmailList {
    param ($xlsFilePath)
    $emails = @()
    try {
        $xls = (Import-Excel -Path $xlsFilePath -NoHeader -EndColumn 1)
    }catch {
        throw "Error reading file $xlsFilePath"
    }
    # Elimina los registros en blanco o null, quita los espacios y lo pone a minusculas
    foreach($item in $xls) {
        if (-not ([string]::IsNullOrWhiteSpace($item.P1))){
            $emails += $item.P1.Trim().toLower()
        }
    }
    [array]$emails
}
# Process a excel files
# and returns a map with the report result
 function ProcessList {
    param ($pathSubscribed)
    $fileNameWithOutExtension = $pathSubscribed | select -expand BaseName  
    $pathSubscribedDirectory = $pathSubscribed.Directory.FullName
    $processingFile = $fileNameWithOutExtension.Substring(0, $fileNameWithOutExtension.Length - $subscribedSuffix.Length)
    $pathUnsubscribed = $pathSubscribedDirectory + [IO.Path]::DirectorySeparatorChar +  $processingFile + $unsubscribedSuffix + '.xlsx'
    [array]$subscribed = GetEmailList $pathSubscribed
    [array]$subscribedUnique = $subscribed | select -Unique
    $subscribedLength = $subscribed.Length
    $subscribedUniqueLength = $subscribedUnique.Length
    $unsubscribedLength = 0
    $mails2Delete = @()
    if ([System.IO.File]::Exists($pathUnsubscribed)) {
        [array]$unsubscribed = GetEmailList $pathUnsubscribed
        $unsubscribedLength = $unsubscribed.Length
        $finalList = $subscribedUnique | Where-Object { (-not($_ -in $unsubscribed )) } | Sort-Object
        $mails2Delete = $subscribedUnique | Where-Object { ($_ -in $unsubscribed) } | Sort-Object
        if (($mails2Delete.Length -ne 0) -or ($subscribed.Length -ne $subscribedUnique.Length)) {
            if ($update) {
                $ExcelParams = @{
                    Path    = $pathSubscribed
                    Show    = $false
                    worksheet = 'Hoja1'
                }
                Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
                $finalList | Export-Excel @ExcelParams
            }
        }
    }else{
        Write-Warning "File $pathUnsubscribed not found!"
    }
    # No return needed
    @{"pathSubscribed" = $processingFile;
        "subscribedDuplicates" = $subscribedLength - $subscribedUniqueLength;
        "subscribedLength" = $subscribedLength; 
        "unsubscribedLength" = $unsubscribedLength;
        "mails2Delete" = $mails2Delete}
}
function GenerateReportExcel {
    param ($summary)
    Write-Output "Summary:"
    # convert the result to list
    $report = @()
    foreach ($key in $summary.Keys){
        #Write-Output "File: ${key}"
        $invalidMails = $summary[$key]["mails2Delete"]  -join ', '
        $summary[$key] | Format-Table -HideTableHeaders
        $row = New-Object PsObject
        $row | Add-Member NoteProperty "List" $summary[$key]["pathSubscribed"]
        $row | Add-Member NoteProperty "Subscribed Mails"  $summary[$key]["subscribedLength"]
        $row | Add-Member NoteProperty "Duplicates"  $summary[$key]["subscribedDuplicates"]
        $row | Add-Member NoteProperty "Unsubscribed Mails" $summary[$key]["unsubscribedLength"]
        $row | Add-Member NoteProperty "Invalid Mail Number" $summary[$key]["mails2Delete"].Length;
        $row | Add-Member NoteProperty "Invalid Mails" $invalidMails
        $report += $row
    }
    #$report = $report | % { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru }
    if (-not ([string]::IsNullOrWhiteSpace($reportXls))){
        Write-Output "Generating report in $reportXls"
        Remove-Item -Path $reportXls -Force -EA Ignore
        $ExcelParams = @{
            Path    = $reportXls
            Show    = $false
            worksheet = 'Report'
            AutoSize = $True
            BoldTopRow = $True
        }
        $report | Sort-Object -Property List |Export-Excel @ExcelParams
    }
}

#main
# Import data
if($null -eq (Get-Module -list ImportExcel)) {
    Install-Module ImportExcel -scope CurrentUser
}
if (-not $update){
    Write-Warning 'Preview mode. No files will be modified.'
    
}
$gSummary = @{}
Write-Output "Processing *$subscribedSuffix.xlsx files in $directory directory..."
$files = Get-ChildItem -Path $directory\*$subscribedSuffix.xlsx -Recurse
if ($files.Length -eq 0){
    Write-Warning "No files found!"
}
else{
    foreach ($file in $files){
        Write-Output "Processing $file..."
        $fileSummary = ProcessList $file
        $gSummary.Add($file, $fileSummary)
    }
}
GenerateReportExcel($gSummary)
Write-Output 'Process finished successfully'