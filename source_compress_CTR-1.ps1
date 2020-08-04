#region script global variables

#get the list of source directories in the original folder
$rootFolder = "C:\hermes\CTR-1\"
$destinFolder = "C:\hermes\CTR-1\"
$archives = "C:\ARC\CTR-1\"
$tempVariable = $rootFolder

#endregion

#region collect information about .txt files

#decide how long back to go
$timespan = new-timespan -Seconds 60

	#create a temporary folder using today's date
	$tempFolderRoot = "$rootFolder"
	$tempFolderDestin = "$destinFolder"
	$tempArchive = "$archives"
	$date = Get-Date
	$date = $date.ToString("yyyy-MM-dd-HH-mm-ss")
	$dateHeader = Get-Date
	$dateHeader = $dateHeader.ToString("yyyy/MM/dd-HH:mm")
	$tempFinalFolder = "$tempFolderRoot$date"
	
	New-Item -ItemType directory -Path $tempFinalFolder
	Write-Output "########################################################################################################################`r`n                         Created on $dateHeader                          `r`n########################################################################################################################`r`n" | Add-Content "$tempFinalFolder.log"
	Write-Output "$date - Temporary workpath $tempFinalFolder created!`r`n" | Add-Content "$tempFinalFolder.log"
	
	#lists files created more than 60 seconds ago
	$inFiles = Get-ChildItem "$rootFolder\*.in" | where {((Get-Date) - $_.LastWriteTime) -gt ($timespan)}
	
	#counts the number of file created more than 60 seconds ago
	$numFiles = ($inFiles.count)
	
	Write-Output "$date - $numFiles files to be processed now!`r`n" | Add-Content "$tempFinalFolder.log"
	Write-Verbose "$numFiles files to be processed now!" -verbose

#endregion

#region temporary workpath for zipping

	#move files to temporary folder before zipping.
	foreach ($in in $inFiles)
	{ 
		Write-Output "$date - File $in found!" | Add-Content "$tempFinalFolder.log"
		Copy-Item "$in" -destination $tempFinalFolder -Force
		Write-Output "$date - File $in moved to temporary workpath $tempFinalFolder`r`n" | Add-Content "$tempFinalFolder.log"
	}
	Start-Sleep -Seconds 1
	
	#Creates zip files to each source directory withing .txt files
	$CompressionToUse = [System.IO.Compression.CompressionLevel]::Optimal
	$IncludeBaseFolder = $false
	$zipTo = "{0}Archive_{1}.zip" -f $tempFolderRoot,$date
		
	#add the files in the temporary location to a zip file
	[Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
	[System.IO.Compression.ZipFile]::CreateFromDirectory($tempFinalFolder, $zipTo, $CompressionToUse, $IncludeBaseFolder)
	Write-Output "`r`n$date - Zip file $zipTo created!" | Add-Content "$tempFinalFolder.log"
	
	#list zip files to move to destination folder
	$sourceZipFiles = Get-ChildItem "$rootFolder\*.zip"
	
	foreach ($zip in $sourceZipFiles)
	{	
		#copy .zip archives to ARC folder
		Copy-Item $zip -destination "$archives"
		Write-Output "`r`n$date - Archive zip file $zip copied to path $archives" | Add-Content "$tempFinalFolder.log"
		
        #move zip files to destination folder
		Move-Item $zip -destination $tempFolderDestin
		Write-Output "`r`n$date - Zip file $zip moved to $tempFolderDestin`r`n" | Add-Content "$tempFinalFolder.log"
	}

#endregion

#region cleaning up files already zipped
	
    #remove files already sent to zip package
	foreach($in in $inFiles)
	{
		Remove-Item "$in"
		Write-Verbose "$in" -verbose
		Write-Output "$date - Collected file $in deleted!" | Add-Content "$tempFinalFolder.log"		
	}
		
	#remove temporary folder on each source dir
	Remove-Item $tempFinalFolder -Recurse
	Write-Output "`r`n$date - Temporary workpath $tempFinalFolder deleted!" | Add-Content "$tempFinalFolder.log"

#endregion

#region check new files for next run

    #list new .txt files to be collected on next run
	$txtFilesNewCount = Get-ChildItem "$rootFolder\*.in"
	$numFilesNewCount = ($txtFilesNewCount.count)
	Write-Output "`r`n$date - $numFilesNewCount new files found to be collected on $tempFolderRoot" | Add-Content "$tempFinalFolder.log"
	Write-Verbose "`r`n$numFilesNewCount new files to be processed on $tempFolderRoot" -verbose
	
	Start-Sleep -Seconds 1

#endregion	

#region update archive zip file

    #update generated zip file with log file
	$log = "$tempFinalFolder.log"
	$zipArc = "{0}Archive_{1}.zip" -f $tempArchive,$date
	Write-Host $zipArc
	[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
	$zipUpdate = [System.IO.Compression.ZipFile]::Open($zipArc,"Update")
	$FileName = [System.IO.Path]::GetFileName($log)
	[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipUpdate,$log,$FileName,"Optimal") | Out-Null
	$zipUpdate.Dispose()
	Remove-Item "$log"

#endregion