#############################################################################
#Josh Woolf - V&V/Blue Team												   	#
#Last Modified - Nov 26, 2018												#
#IAVM, STIG, CCRI Scoring										   						#
#XML to CSV -> CSV Combiner -> STIG Scoring All-in-One Tool				  	#
#############################################################################


##############check to see if STIG_Score.xlam/IAVM_Score.xlam is present in the Addins Folder##############
robocopy /xc /xo "\\usr.osd.mil\org\OSD\JSP\JP2\JP22\JP221\V&V\Tools\STIGscore\" "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns\" STIG_Score.xlam /NFL /NDL /NJH /NJS /nc /ns /np
robocopy /xc /xo "\\usr.osd.mil\org\OSD\JSP\JP2\JP22\JP221\V&V\Tools\IAVMscore\" "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns\" IAVM_Score.xlam /NFL /NDL /NJH /NJS /nc /ns /np
#/xn
###Variables###
$FinalSTIGscore = 0
$rootPath = $pwd | split-path -leaf
$IAVM_File = Get-ChildItem -filter vulns.csv
$Prompt = Read-host "Choose one of the following options (1,2,3) based off of assessment requirements: 
	1. IAVMs only 
	2. Full Assessment 
	3. Exit
Option"

###Switch Statement to determine working requirements (IAVMs only/full assessment)###

do {

Switch($Prompt)
{	#IAVMs Only
	1 { 
		##############Scoring IAVMs from CSV##############
		Write-Output "Scoring IAVMs... `n"

		#if vulns.csv doesn't exist, say so and exit
		if ($IAVM_File.length -eq 0)
		{Write-Host "vulns.csv not found; save the IAVMs to the root folder as 'vulns.csv' and rerun the tool." 
		Write-Host "Press Any Key To Exit" 
		$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
		exit  }
		
		
		Get-ChildItem -filter vulns.csv |
		Foreach-Object {
			$excel = New-Object -comobject Excel.Application

			#open file
			$FilePath = $_.FullName
			$workbook = $excel.Workbooks.Open($FilePath)
			$worksheet = $workbook.worksheets.item(1)
			$sheetName = $worksheet.name

			#access the Application object and run a macro
			$excel.Run("'C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns\IAVM_Score.xlam'!IAVMscore")

			$worksheet = $workbook.Sheets.Item(2)
			$worksheet.columns.item('b').NumberFormat = "@"
			
			
			#pull values from the excel 'output' sheet to display here
			#####IPs Scanned######
			$totalIPs = $worksheet.cells.item(1,2).Text
			#####Open Total#####
			$IAVMhighs = $worksheet.cells.item(4,2).Text
			$IAVMmediums = $worksheet.cells.item(5,2).Text
			$IAVMLows = $worksheet.cells.item(6,2).Text
			#####Uniques#####
			$uHighs = $worksheet.cells.item(8,2).Text
			$uMediums = $worksheet.cells.item(9,2).Text
			$uLows = $worksheet.cells.item(10,2).Text
			#####Score#####
			$IAVMscore = $worksheet.cells.item(16,2).Text
			
			#write the output here, whatever you want
			Write-Output "Findings for $rootPath IAVMs"
			Write-Output "-------------------------------------------------------------------"
			Write-Output "Total Critical/High Findings: $IAVMhighs, Total Medium Findings: $IAVMmediums, Total Low Findings: $IAVMlows"
			Write-Output "Unique Critical/High Findings: $uHighs, Unique Medium Findings: $uMediums, Unique Low Findings: $uLows `n"
			Write-Output "$rootPath IAVM Score: $IAVMscore"
			Write-Output "------------------------------------------------------------------- `n"

			#Clean up task manager
			$workbook.Close($false)
			$excel.quit()
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
			[GC]::Collect()
			
			} | Tee-Object -file "$pwd\results.txt"
						
			Get-ChildItem -filter vulns.csv | Rename-Item -NewName "$rootPath IAVMs.csv"
			Write-Output "These results can be found in the results.txt file that was created in the $rootPath folder. `n `n"
						
			#exit PS
			Write-Host "Press Any Key To Exit" 
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
			exit 
	}
	
	#Full Assessment
	2 {
		##############Scoring IAVMs from CSV##############
		Write-Output "Scoring IAVMs... `n"
		if ($IAVM_File.length -eq 0) {
		Write-Host "vulns.csv not found; save the IAVMs to the root folder as 'vulns.csv' and rerun the tool.  Continuing to evaluate STIGs." }
		
		Get-ChildItem -filter vulns.csv |
		Foreach-Object {
			$excel = New-Object -comobject Excel.Application

			#open file
			$FilePath = $_.FullName
			$workbook = $excel.Workbooks.Open($FilePath)
			$worksheet = $workbook.worksheets.item(1)
			$sheetName = $worksheet.name

			#access the Application object and run a macro
			$excel.Run("'C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns\IAVM_Score.xlam'!IAVMscore")

			$worksheet = $workbook.Sheets.Item(2)
			$worksheet.columns.item('b').NumberFormat = "@"
			
			
			#pull values from the excel 'output' sheet to display here
			#####IPs Scanned######
			$totalIPs = $worksheet.cells.item(1,2).Text
			#####Open Total#####
			$IAVMhighs = $worksheet.cells.item(4,2).Text
			$IAVMmediums = $worksheet.cells.item(5,2).Text
			$IAVMLows = $worksheet.cells.item(6,2).Text
			#####Uniques#####
			$uHighs = $worksheet.cells.item(8,2).Text
			$uMediums = $worksheet.cells.item(9,2).Text
			$uLows = $worksheet.cells.item(10,2).Text
			#####Score#####
			$IAVMscore = $worksheet.cells.item(16,2).Text
			
			#write the output here, whatever you want
			Write-Output "Findings for $rootPath IAVMs"
			Write-Output "-------------------------------------------------------------------"
			Write-Output "Total Critical/High Findings: $IAVMhighs, Total Medium Findings: $IAVMmediums, Total Low Findings: $IAVMlows"
			Write-Output "Unique Critical/High Findings: $uHighs, Unique Medium Findings: $uMediums, Unique Low Findings: $uLows `n"
			Write-Output "$rootPath IAVM Score: $IAVMscore"
			Write-Output "------------------------------------------------------------------- `n"

			#Clean up task manager
			$workbook.Close($false)
			$excel.quit()
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
			[GC]::Collect()
			
			} | Tee-Object -file "$pwd\results.txt"

		Get-ChildItem -filter vulns.csv | Rename-Item -NewName "$rootPath IAVMs.csv"

		###################XML to CSV###################

		#Testing pulling each individual attribute out of XML and assigning it to a variable
			#for all STIG types (folders) in System (root) folder
			
			
			$fileCount = ((Get-ChildItem -Recurse -filter *xccdf-res.xml | Measure-Object).Count)
			Write-Progress -Activity Converting -Status 'Progress->' 
			#progress bar increment
			$i = 0
			
			
			#create the CSVs
			Get-ChildItem -Recurse -filter *xccdf-res.xml | 
			Foreach-Object {
			[xml]$inputFile = Get-Content $_.FullName
			$path = (get-childitem $_.FullName | split-path )
			$lastPath = $path | split-path -leaf
			Write-Output "Generating CSV from $_ XML data in $lastPath."
			set-location -Path $path
			$Rename = $inputFile.TestResult.{target-address}
			
				#pull fields from XML into a csv
				$inputFile.TestResult.{rule-result} | 
									Select-Object @{ expression={$_.idref}; label="RuleID" }, 
												  @{ expression={$_.result}; label="Status" },
												  @{ expression={$_.severity}; label="Severity" } |
												  ConvertTo-Csv -NoTypeInformation | Set-Content -Path ('.\' + $_.Basename + '.csv') -Encoding UTF8
								
				 #Trimming the Rule ID to something legible, and Status to reflect ckl statuses, and spitting it back into the csv 
				 $trimRule = Import-Csv ('.\' + $_.Basename + '.csv')
				 $trimRule | ForEach-Object { 
					$_.RuleID = ($_.RuleID -replace ".*S", "S" -replace "_.*") 
					$_.Status = ($_.Status -replace "pass", "Not A Finding" -replace "fail", "Open" -replace "error", "Not A Finding") 
					} 
					$trimRule | export-csv ('.\' + $_.Basename + '.csv') -notype
					
					#renames the csv to reflect the IP address
					Rename-Item -Path ('.\' + $_.Basename + '.csv') -NewName "$($Rename).csv"
					
		###################CSV Combiner###################
		#creates the "output" folder if it doesnt exist for each STIG type
		New-Item -ItemType Directory -Force -Path $pwd\output | Out-Null


		#combine the csvs, unique IPs can be determined via the newly added "FileName" column

		Get-ChildItem $pwd\*.csv -PipelineVariable File |
		  ForEach-Object { Import-Csv $_ | Select *,@{l='FileName';e={$File.Name}}} |
		  Export-Csv ('.\output\' + (split-path -path $path -Leaf)  + '.csv') -NoTypeInformation
		  
		  Copy-Item ('.\output\' + (split-path -path $path -Leaf)  + '.csv') -Destination "..\"
		  Remove-Item .\output\ -Force -Recurse
		  
		  $i++
		  Write-Progress -activity "Converting" -Status "Converted $i of $fileCount" -percentComplete (($i / $fileCount) * 100)
		  	  
		  set-location ..
		} #end of the foreach loop
		#remove progress bar
			Write-Progress -Completed -Activity "progress bar, no progressing!"

		#done merging 
		Write-Output "CSV merge complete. `n"
		Write-Output "Scoring STIGs... `n"
		###################Run Combined CSV Through Excel Scoring Macro###################

		Get-ChildItem $_.FullName -exclude *.ps1, *.docx, *.txt, *.ckl, *.zip | Where-Object{!($_.PSIsContainer) -and $_.Name -notlike "*IAVMs*"} |
		Foreach-Object {
			$ErrorActionPreference= 'silentlycontinue'
			
			$excel = New-Object -comobject Excel.Application

			#open file
			$FilePath = $_.FullName
			$workbook = $excel.Workbooks.Open($FilePath)
			$worksheet = $workbook.worksheets.item(1)
			$sheetName = $worksheet.name

			#access the Application object and run a macro
			$excel.Run("'C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns\STIG_Score.xlam'!STIGscore")

			#access the "output" sheet created by the macro which contains the information the message box popup lists
			$worksheet = $workbook.Sheets.Item(5)
			$worksheet.columns.item('b').NumberFormat = "@"
			
			
			#pull values from the excel 'output' sheet to display here
			#####IPs Scanned######
			$totalIPs = $worksheet.cells.item(1,2).Text
			#####Open Total#####
			$openHighs = $worksheet.cells.item(4,2).Text
			$openMediums = $worksheet.cells.item(5,2).Text
			$openLows = $worksheet.cells.item(6,2).Text
			#####Uniques#####
			$uniqueHighs = $worksheet.cells.item(8,2).Text
			$uniqueMediums = $worksheet.cells.item(9,2).Text
			$uniqueLows = $worksheet.cells.item(10,2).Text
			#####Total Possible#####
			$totalHighs = $worksheet.cells.item(12,2).Text
			$totalMediums = $worksheet.cells.item(13,2).Text
			$totalLows = $worksheet.cells.item(14,2).Text
			

			#set unique counts = total counts if only 1 IP
			if($totalIPs -eq 1) {
				$uniqueHighs = $openHighs
				$uniqueMediums = $openMediums
				$uniqueLows = $openLows
				}
			
			#pull STIG Score
			$STIGScore = $worksheet.cells.item(16,2).Text
		 
			#assign STIG Weighted Value for CCRI Score integration based on threshold
			if(($STIGScore -as [double]) -ge 20.00) {
			$weightedValue = 1.0
			$Concern = "Critical Concern"}
			elseif(($STIGScore -as [double]) -ge 10.00 -AND ($STIGScore -as [double]) -le 20.00) {
			$weightedValue = 0.40
			$Concern = "Moderate Concern"}
			else {
			$weightedValue = 0.0
			$Concern = "Minimal Concern"}
			
			#output the Score and CCRI Modifier
			Write-Output "Findings for $sheetName"
			Write-Output "-------------------------------------------------------------------"
			Write-Output "Total Open CAT Is: $openHighs, Total Open CAT IIs: $openMediums, Total Open CAT IIIs: $openLows"
			Write-Output "Unique CAT Is: $uniqueHighs, Unique CAT IIs: $uniqueMediums, Unique Lows: $uniqueLows"
			Write-Output "Total Possible CAT Is: $totalHighs, Total Possible Mediums: $totalMediums, Total Possible Lows: $totalLows `n"
			Write-Output "$sheetName STIG Score: $STIGScore% $Concern"
			Write-Output "CCRI Score Modifier: $weightedValue"
			Write-Output "------------------------------------------------------------------- `n"

			
			
			#add the weighted value for this STIG to the cumulative STIG score to be included in the overall CCRI Score
			$FinalSTIGscore = $FinalSTIGscore + $weightedValue

			#Clean up task manager
			$workbook.Close($false)
			$excel.quit()
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
			[GC]::Collect()
			
			} | Tee-Object -file "$pwd\results.txt" -append
			
			Write-Output "The Overall STIG Score Modifier is: $FinalSTIGscore `n" | Tee-Object -file "$pwd\results.txt" -append
			
			#Calculating CCRI Score
			[double]$CCRIscore = [double]$IAVMscore + [double]$FinalSTIGscore
			Write-output "**The CCRI Score is: $CCRIscore.** `n" | Tee-Object -file "$pwd\results.txt" -append
						
			Write-Output "These results can be found in the results.txt file that was created in the $rootPath folder. `n `n"
			
			#exit PS
			Write-Host "Press Any Key To Exit" 
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
			exit 
	}
	#Exit
	3 { exit }
	Default {continue} 
#end switch	
}
#end do
} while($Prompt -notmatch "[123]")
