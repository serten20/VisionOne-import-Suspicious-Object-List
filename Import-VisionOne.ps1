$iocs = Get-Content -Path .\suspicious-object.txt
$i = 0

$ruta = $PWD

$ficheroExcel = "$ruta\VisionOneFormat.xlsx"
$csv = "$ruta\VisionOneImport.csv"

write-host "--------------------------------------------------------------"
Write-host "------------------------ Vision One --------------------------"
write-host "--------------------------------------------------------------"
write-host ""



 $Description = read-host "Write a description: "
 
 $ipv4 = @"
^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$
"@

$ipv6 = @"
^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$
"@

$url = @"
(?i)\b((?:[a-z][\w-]+:(?:\/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}\/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))
"@

 $domain = @"
^(?!-)(?:(?:[a-zA-Zd][a-zA-Zd-]{0,61})?[a-zA-Zd].){1,126}(?!d+)[a-zA-Zd]{0,63}$
"@

$email = @"
[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?
"@




$Excel = New-Object -com Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false


$Book = $Excel.Workbooks.add()
$WorkSheet = $Book.Sheets.Item(1)
$WorkSheet.Name = "VisionOne"


ForEach($ws in $WorkSheet)

{

    $ws.Range($ws.Rows(1),$ws.Rows($ws.Rows.Count)).Delete() | Out-null

}




$WorkSheet.Select()
$WorkSheet.Cells.Item(1,1) = "Type"
$WorkSheet.Cells.Item(1,2) = "Object" 
$WorkSheet.Cells.Item(1,3) = "Description" 

$row = 1					
   
        foreach ($ioc in $iocs)
            {
				
				$i++
				#$myvar += $ioc
					
					
	
				
				if ($ioc.Length -eq 32){
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Cyan -NoNewLine
					Write-Host " (Hash " -NoNewLine
					Write-Host "md5" -ForegroundColor Cyan -NoNewLine
					Write-Host ") " -NoNewLine
					Write-Host " Not supported in VisionOne, will not append to csv file"
				}

				elseif ($ioc.Length -eq 40) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Magenta -NoNewLine
					Write-Host " (Hash " -NoNewLine
					Write-Host "Sha1" -ForegroundColor Magenta -NoNewLine
					Write-Host ")"
				
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "sha1"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"

					
				}

				elseif ($ioc.Length -eq 64) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Green -NoNewLine
					Write-Host " (Hash " -NoNewLine
					Write-Host "Sha256" -ForegroundColor Green -NoNewLine
					Write-Host ")"
					
					
					#$row = $incre+1 
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "sha256"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
					
					
				
				}
				
				elseif ($ioc -match $ipv4) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Yellow -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "IPv4" -ForegroundColor Yellow -NoNewLine
					Write-Host ")" 
					
					
				
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "ip"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
					
				
				}
				
				elseif ($ioc -match $ipv6) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Yellow -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "IPv6" -ForegroundColor Yellow -NoNewLine
					Write-Host ")" 
								
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "ip"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
					
				
				}
				
				elseif ($ioc -match $url) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Red -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "URL" -ForegroundColor Red -NoNewLine
					Write-Host ")" 
					
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "url"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
					
				
				}
				
				elseif ($ioc -match $email) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Gray -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "Email" -ForegroundColor Gray -NoNewLine
					Write-Host ")" 
									
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "email_sender"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
				
				}
				
				elseif ($ioc -match $domain) {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Magenta -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "Domain" -ForegroundColor Magenta -NoNewLine
					Write-Host ")" 
									
					
					$row++ | Out-null
					$column = 1
					$WorkSheet.Cells.Item($row,$column) = "domain"
					
					$row | Out-null
					$column = 2
					$WorkSheet.Cells.Item($row,$column)  = "$ioc"
					
					$row | Out-null
					$column = 3
					$WorkSheet.Cells.Item($row,$column)  = "$Description"
				
				}
				

				else {
					Write-Host "   [$i]" -NoNewLine
					Write-Host " $ioc" -ForegroundColor Red -NoNewLine
					Write-Host " (" -NoNewLine
					Write-Host "Other" -ForegroundColor Red -NoNewLine
					Write-Host " Not supported in VisionOne, will not append to csv file"
					
					

					
					
				}
					
                    
             
            }

		
		
		
		$Book.SaveAs($ficheroExcel) 
		
		$Book.Close()
		$Excel.Quit()
		
		
		


		Write-Host " The csv file is being generated, wait a few seconds. (" -nonewline 
		write-host "$csv" -nonewline -ForegroundColor Green
		write-host ")"
		
		Start-Sleep -seconds 10
		
		$wb = $Excel.Workbooks.Open($ficheroExcel)

        foreach ($ws in $wb.Worksheets)

        {

            
			$ws.SaveAs($csv, 6)

        }

        $Excel.Quit()
		
		remove-item $ficheroExcel
		


pause