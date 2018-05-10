Function IndexLinks {
Param ($path)

$folder = $env:TEMP + "\" +  [System.IO.Path]::GetFileNameWithoutExtension($path)
$zippath = $env:TEMP + "\" +  [System.IO.Path]::GetFileNameWithoutExtension($path) + '.zip'
$Filename = [System.IO.Path]::GetFileName($path)
Write-Host 'Checking' $Filename 'for links'
Copy-Item $path -Destination $zippath
Expand-Archive $zippath -DestinationPath $folder -Force


$slides = get-childitem $folder -filter *.xml.rels -recurse
$xml = New-Object -TypeName XML
$Links = @()
ForEach ($slide in $slides) {
    
    $xml.load($slide.fullname)
    #Write-Host $slide.FullName
    #$xml.Relationships.ChildNodes | format-list -property $_.Target
    $LinkNodes = $xml.Relationships.ChildNodes | where {$_.Target -like "file*"} 
    foreach ($LinkNode in $LinkNodes){
        $SlideNumber = $slide.Name -replace '.xml.rels',''
        $HyperLink = $LinkNode.Target -replace 'File:///','' -replace '%20',' '
        $Link = new-object psobject -prop @{Filename=$Filename;Path=$path; Hyperlink_Type='Hyperlink'; Link=$Hyperlink; Slide_Number=$SlideNumber; New_link=''; ID=$LinkNode.Id}
        $Links += $Link   
    }
    
    #clear
    #$xml.Save($slide.FullName)
    
} #End of main loop
$Links = $Links | Sort-Object -Property Slide_Number

Remove-Item $zippath
Remove-Item $folder -Recurse

Return $Links
}

Function FixLinks {
    param ($Hyperlink)
    if ($hyperlink.IndexOf('#') -gt 5){
        Write-host 'powerpoint link contains bookmark'
        $index = $hyperlink.IndexOf('#')
        $Hyperlink = $Hyperlink.Substring(0, $index)
    }
    if ([System.IO.Path]::GetExtension($hyperlink) -like ".lnk"){$Hyperlink = $Hyperlink -replace '.lnk',''}
    $SearchFolder = $Directory.indexofany("\",(([system.io.path]::GetPathRoot($directory).length) + 1 )) 
    $SearchFolder = $Directory.substring(0,$SearchFolder)
    $FileName = [System.IO.Path]::GetFileName($Hyperlink)
    if (($hyperlink -like "\\srv*") -and ($hyperlink -notlike $directory.Substring(0,9))){
        $NewHyperlink = (Get-ChildItem -path $SearchFolder -Filter $FileName -Exclude "*.lnk" -Recurse).FullName | Select-Object -First 1
        $SearchFolder = [system.io.path]::GetPathRoot($Directory)
        if ($NewHyperLink -like ''){$NewHyperlink = (Get-ChildItem -path $SearchFolder -Filter $FileName -Exclude "*.lnk" -Recurse).FullName | Select-Object -First 1}
        Return $NewHyperlink;
        break
       }
    If (Test-Path $Hyperlink){
        Return 'Functional Link'
    }
    Else {
        #$NewHyperlink = (Get-ChildItem -path $Directory -Filter $FileName -Recurse).FullName | Select-Object -First 1
        $NewHyperlink = (Get-ChildItem -path $SearchFolder -Filter $FileName -Exclude "*.lnk" -Recurse).FullName | Select-Object -First 1
        $SearchFolder = [system.io.path]::GetPathRoot($Directory)
        if ($NewHyperLink -like ''){$NewHyperlink = (Get-ChildItem -path $SearchFolder -Filter $FileName -Exclude "*.lnk" -Recurse).FullName | Select-Object -First 1}
        Return $NewHyperlink
    }
    Return ''
}

$OutputPath = $env:USERPROFILE + '\Desktop\Powerpoint_log.csv'
$links = @()
clear
$Directory = Read-Host "Where are the powerpoints?"
clear
Write-Host "Finding Powerpoints"
$Powerpoints = Get-ChildItem -Path $Directory -Filter *.ppt* -recurse -Exclude '*.ppt','*.pps','*.lnk'
Write-Host 'Powerpoints found, checking powerpoints for Links'
ForEach ($Powerpoint in $powerpoints){
    $links += IndexLinks $Powerpoint.FullName    
}
Write-Host 'Links found, Checking if links are working'
ForEach ($link in $links){
    Write-Host "Testing links in" $link.Filename
    $Link.New_Link = FixLinks $link.Link
}
$Links | select-object Filename, Slide_Number, Link, New_link, Path, ID | Export-csv -path $OutputPath -encoding ascii -NoTypeInformation -Force
Write-host 'Finished'