$sheetsIco = "C:\Program Files\LinkIconSetter\sheets.ico"
$docsIco = "C:\Program Files\LinkIconSetter\docs.ico"

Get-ChildItem "C:\Users\Reubin.DESKTOP-I0L3FOR\Desktop\*.*" -include *.url | 
Foreach-Object {
    #$content = Get-Content $_.FullName
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

        if (-Not $iconLocation -match $docsIco)
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