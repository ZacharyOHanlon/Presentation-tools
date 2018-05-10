Function ReplaceLinks {
Param ($Links)
$path = $Links[0].Path
$folder = $env:TEMP + "\" +  [System.IO.Path]::GetFileNameWithoutExtension($path)
$zippath = $env:TEMP + "\" +  [System.IO.Path]::GetFileNameWithoutExtension($path) + '.zip'
$Filename = [System.IO.Path]::GetFileName($path)
Write-Host 'Replacing links in ' $Filename
Copy-Item $path -Destination $zippath
Expand-Archive $zippath -DestinationPath $folder -Force

    ForEach ($link in $links){
    $Link.Link = 'file:///' + $link.Link
    $Link.New_Link = 'file:///' + $link.New_Link
    $Link.Link = $link.Link -replace ' ','%20'
    $Link.New_link = $link.New_Link -replace ' ','%20'
    $filter = $Link.Slide_Number + '.xml.rels'
    $slide = get-childitem $folder -filter $filter -recurse | Select-Object -First 1
    $xml = New-Object -TypeName XML
    $xml.Load($slide.FullName)
    

    $LinkNode = $xml.SelectNodes("//*[@Id]") | Where-Object {$_.Id -like $Link.Id}
    $LinkNode
    if ($linkNode.Target -like $Link.link){$LinkNode.Target = $link.New_Link}
    $xml.save($slide.FullName)
    }
$Extension = [System.IO.Path]::GetExtension($path)
$Fixed = '-fixed' + $Extension
$OutputPath = $path -replace $Extension,$Fixed
$contents = $folder + '\*'
Compress-Archive $contents -DestinationPath $zippath -Force
Copy-Item $zippath -Destination $OutputPath
Remove-Item $zippath
Remove-Item $folder -Recurse


}



clear
$LinkPath = $env:USERPROFILE + '\Desktop\Powerpoint_log.csv'
$Logs = import-csv $LinkPath
$logs = $logs | Sort-Object -Property Path -Descending | Where-Object {(($_.New_Link -ne 'Functional Link') -and ($_.New_Link -ne ""))}
$Files = $logs | sort-object -Property Path -Unique
ForEach ($File in $Files){
    $links = $logs | Where-Object {$_.path -like $file.path}
   ReplaceLinks $links
}
Write-Host "Finished replacing links"