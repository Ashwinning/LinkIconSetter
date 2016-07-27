$sheetsIco = "$pwd\sheets.ico"
$docsIco = "$pwd\docs.ico"

Get-ChildItem "$env:USERPROFILE\Desktop\*.*" -include *.url | 
Foreach-Object {
    $linkItem = $_.FullName
    $sh = New-Object -COM WScript.Shell
    $objShortcut = $sh.CreateShortcut($linkItem)
    if ($objShortcut.TargetPath -match ‘google.com/spreadsheets’)
    {
        #this shortcut's currently set icon location
        $iconLocation = GetIconLocation -file $objShortcut.FullName
        If([string]::IsNullOrEmpty($iconLocation))
        {
            AddIconData -file $objShortcut.FullName -icon $sheetsIco
        }

        ElseIf (-Not $iconLocation.Contains($sheetsIco))
        {
            RemoveIconData -file $objShortcut.FullName
            AddIconData -file $objShortcut.FullName -icon $sheetsIco
        }

    }
    if ($objShortcut.TargetPath -match ‘google.com/document’)
    {
        #this shortcut's currently set icon location
        $iconLocation = GetIconLocation -file $objShortcut.FullName
        If([string]::IsNullOrEmpty($iconLocation))
        {
            AddIconData -file $objShortcut.FullName -icon $docsIco
        }

        ElseIf (-Not $iconLocation.Contains($docsIco))
        {
            RemoveIconData -file $objShortcut.FullName
            AddIconData -file $objShortcut.FullName -icon $docsIco
        }
    }
    $objShortcut.Save()
}

function RemoveIconData
{
    Param ([string] $file)
    (Get-Content $file) -notmatch "IconIndex" | Out-File $file
    (Get-Content $file) -notmatch "IconFile" | Out-File $file
}

function AddIconData
{
    Param ([string] $file, [string] $icon)
    #Add IconIndex
    Add-Content $file "IconIndex=0"
    #Add Icon Location
    Add-Content $file "IconFile=$icon"
}

function GetIconLocation
{
    Param([string] $file)
    Get-Content $file | Where-Object { $_.Contains("IconFile=") }
}