.\Invoke-Kape.ps1

$memory_data_directory = 'Insert file path' 
$Kape_destination_directory = 'Insert output file path' 
if (-not (Test-Path -Path $kape_destination_directory -PathType Container)) {
#Directory doesnt exist, create it
New-Item -Path $kape_destination_directory -ItemType Directory -Force
Write-Host "Directory created at $kape_destination_directory" 
} else {
Write-Host "Directory already exists"
}
get-childitem -Directory $memory_data_directory).name | ForEach-Object {
    Invoke-Kape  -msource $memory_data_directory\$_ -mdest $kape_destination_directory\$_ -module DumpIt_Memory,Volatility_amcache,Volatility_clipboard,Volatility_cmdline,Volatility_cmdscan,Volatility_connections,Volatility_connscan,Volatility_consoles,Volatility_dlllist,Volatility_driverirp,Volatility_hollowfind,Volatility_idt,Volatility_malfind,Volatility_modscan,Volatility_modules,Volatility_netscan,Volatility_notepad,Volatility_pslist,Volatility_psscan,Volatility_pstree,Volatility_psxview,Volatility_shimcache,Volatility_sockets,Volatility_sockscan,Volatility_ssdt,Volatility_userassist,Volatility_userhandles --mef csv
}

  $kape_destination_directory = 'Insert file path' 

  $ExcelObject=New-Object -ComObject excel.application
    $ExcelObject.visible=$false 
   # $excel.DisplayAlerts = $false
    $ExcelFiles=Get-ChildItem -Path $kape_destination_directory\$_ -Recurse -Include *.csv | Where-Object { $_.Length -gt 1024 }

    $Workbook=$ExcelObject.Workbooks.add()
    $Worksheet = $Workbook.Sheets.Item(1)
    

    foreach($ExcelFile in $ExcelFiles){
 
        $Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)
        $Everysheet=$Everyexcel.sheets.item(1)
        $Everysheet.Copy($Worksheet)
    $Everyexcel.Close()
 
    }
$Workbook.SaveAs("$kape_destination_directory\KAPE-MEMORY-OUTPUT.csv")
$ExcelObject.Quit()
