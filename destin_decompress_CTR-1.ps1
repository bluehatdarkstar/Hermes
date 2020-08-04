#region script global variables

    #get the root directory to extract ".in" files
    $rootFolder = "c:\hermes\CTR-1\"
    #get the path of logs
    $logFolder = "$rootFolder\logs\"
    #build date with 2 different formats
    $date = Get-Date
    $date = $date.ToString("yyyy-MM-dd-HH-mm-ss")
    $dateHeader = Get-Date
    $dateHeader = $dateHeader.ToString("yyyy/MM/dd-HH:mm")

#endregion

#region select zip packages and create log folder

    $destinZipFiles = Get-ChildItem "$rootFolder\*.zip"

    foreach ($zipped in $destinZipFiles)
	    {	

            if (!(Test-Path -path $logFolder))
            {
                New-Item $logFolder -Type Directory
            }
		
        $zipped = Split-Path -Path $zipped -leaf
        Write-Host $zipped
	
        Start-Sleep -Seconds 1	
	    }

#endregion

#region search for ".in" files in zip packages
		
		$search = "in"
		$zips = "$rootFolder"
		$Manifest = "$logFolder\$date.log"
        
        Write-Output "########################################################################################################################`r`n                         Log file created on $dateHeader                          `r`n########################################################################################################################`r`n" >> $Manifest
 
    Function GetZipFileItems
    {
		Param([string]$zip)
		$split = $split.Split(".")
		$shell = New-Object -Com Shell.Application
		$zipItem = $shell.NameSpace($zip)
		$items = $zipItem.Items()
        $split = $zipFile.Split("\")[-1]
        
        Write-output "Contents of $split`r`n"
        GetZipFileItemsRecursive $items
    }

    Function GetZipFileItemsRecursive
    {
      Param([object]$items)
      ForEach($item In $items)
		{
			
			$strItem = [string]$item.Name
			If ($strItem -Like "*$search*")
			{
				If ((Test-Path ($zips + "\" + $strItem)) -eq $False)
					{
						$zipFile = Split-Path -Path $zipFile -leaf
						Write-Host "Copied file : $strItem from zip-file $zipFile to destination folder"
						$shell.NameSpace($zips).CopyHere($item)
					}
					
				Write-output "$strItem"
			}
		}
    }

    Function GetZipFiles
	{
		$zipFiles = Get-ChildItem -Path $zips -Recurse -Filter "*.zip" | % { $_.DirectoryName + "\$_" }
		ForEach ($zipFile in $zipFiles)
		{
			$split = $zipFile.Split("\")[-1]
            GetZipFileItems $zipFile
			$count = GetZipFileItems $zipFile      
            $ZipFilesItemsCount = ($count.count)-1
            Write-Output "`r`n$ZipFilesItemsCount files found on $Split`r`n"
			
		}

    }

#endregion

#region counting of ".in" files and outputs to the log file

    GetZipFiles >> $Manifest
    $FileContent = Get-Content $Manifest
	$Matches = Select-String -InputObject $FileContent -Pattern '.in' -AllMatches
    $total = $Matches.Matches.Count
    
	write-output "`r`nTotal files found in zip files: $total`r`n" >> $Manifest
    
	$zippedFile = Split-Path $zipped -leaf
       
	Write-Output "All .in files were extracted!`r`n" >> $Manifest

#endregion

#region remove .zip packages after extraction and outputs to log file

    foreach ($zipped in $destinZipFiles)
	    {	
            #delete zipfile
	        Remove-Item "$zipped"
            $zipped = Split-Path $zipped -leaf
	        Write-Output "Zip file $zipped was deleted!" >> $Manifest
        }

#endregion