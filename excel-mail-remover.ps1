param (
    [string]$unsubscribedSuffix = '_unsuscribed_members',
    [string]$subscribedSuffix = '_subscribed_members',
    [string]$directory = $PSScriptRoot,
    [switch]$update = $false
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

 function ProcessList {
    param ($pathSubscribed)
    Write-Output "Processing $pathSubscribed..."
    $fileNameWithOutExtension = $pathSubscribed | select -expand BaseName  
    $pathUnsubscribed = $directory + '\' +  $fileNameWithOutExtension.Substring(0, $fileNameWithOutExtension.Length - $subscribedSuffix.Length) + $unsubscribedSuffix + '.xlsx'
    if ([System.IO.File]::Exists($pathUnsubscribed)) {
        Write-Output "Getting subscribed emails $pathSubscribed ..."
        [array]$subscribed = GetEmailList $pathSubscribed
        Write-Output "Getting unsubscribed emails $pathUnsubscribed ..."
        [array]$unsubscribed = GetEmailList $pathUnsubscribed
        $finalList = $subscribed | Where-Object { (-not($_ -in $unsubscribed )) } 
        $mails2Delete = $subscribed | Where-Object { ($_ -in $unsubscribed) } 
        if ($mails2Delete.Length -ne 0){
            Write-Output [string]::Format("{0} e-mails to delete from file...", $mails2Delete.Length)
            Write-Output "---------------------------------"
            foreach ($item in $mails2Delete) {
                [string]::Format("'{0}'", $item)
            }
            if ($update) {
                Write-Output "Updating $pathSubscribed file..."
                $ExcelParams = @{
                    Path    = $pathSubscribed
                    Show    = $false
                    Verbose = $true
                    worksheet = 'Hoja1'
                }
                Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
                $finalList | Export-Excel @ExcelParams
            }
        }
        else
        {
            Write-Output "No unsubscribed email found."  
        }
    }else{
        Write-Output "Unsubscribed file $Unsubscribed not found. $pathSubscribed won't be modified."
    }
}

#main
# Import data
if($null -eq (Get-Module -list ImportExcel)) {
    Install-Module ImportExcel -scope CurrentUser
}
if (-not $update){
    Write-Output 'Preview mode. No files will be modified.'
    
}
"Processing *$subscribedSuffix.xlsx files in $directory directory..."
$files = Get-ChildItem -Path $directory\*$subscribedSuffix.xlsx
if ($files.Length -eq 0){
    Write-Output "No files found!"
}
else{
    foreach ($file in $files){
        ""
        ProcessList $file
    }
}
Write-Output "Process finished successfully"


