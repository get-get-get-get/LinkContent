function Get-LinkContent {
    <#
    .SYNOPSIS
    Parses content of .lnk file
    .DESCRIPTION
    'Get-LinkContent' parses link/shortcut objects
    .PARAMETER Path
    Path to .lnk object
    .EXAMPLE
    Get-LinkContent -Path C:\Users\hector\AppData\Roaming\Microsoft\Windows\Recent\Restore.lnk
    .EXAMPLE
    $Path = "C:\Users\hector\AppData\Roaming\Microsoft\Windows\Recent\Restore.lnk"
    Get-LinkContent $Path
    .INPUTS
    String
    Accepts paths to link/shortcut objects
    .OUTPUTS
    ComObject
    A Shortcut ComObject representing the link supplied
    #>
    
    [OutputType('System.__ComObject#{f935dc23-1cf0-11d0-adb9-00c04fd58a0b}')]
    Param (
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name', 'File', 'FilePath')]
        [ValidateScript({ Test-Path $_ })]
        [String[]]
        $Path
    )

    ForEach ($File in $Path) {
        $FullPath = (Resolve-Path $File).Path      
        (New-Object -ComObject WScript.Shell).CreateShortcut($FullPath)
    }
        
}


function Get-DirectoryLinkContent {
    <#
    .SYNOPSIS
    Parses contents of all .lnk files within directory
    .DESCRIPTION
    Uses 'Get-LinkContent' to parse all link/shortcut objects within directory
    .PARAMETER Path
    Path to .lnk object
    .PARAMETER Recurse
    Get link content recursively
    .EXAMPLE
    Get-DirectoryLinkContent C:\Users\hector\ -Recurse
    .EXAMPLE
    $Path = "C:\Users\hector\AppData\Roaming\Microsoft\Windows\Recent\Restore.lnk"
    Get-LinkContent -Path $Path
    .INPUTS
    String
    Accepts paths to directories
    .OUTPUTS
    ComObject
    An array of Shortcut ComObjects representing the links within directory
    #>
    [OutputType('System.__ComObject#{f935dc23-1cf0-11d0-adb9-00c04fd58a0b}')]
    Param (
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name', 'Directory', 'FilePath')]
        [ValidateScript({ Test-Path $_ })]
        [String[]]
        $Path,

        [Parameter(Position = 1)]
        [switch]
        $Recurse
    )

    # TODO: how to do more elegantly (or, how to just accept common parameters?)
    if ($Recurse -eq $False) {
        $Links = Get-ChildItem $Path -Force -Filter *.lnk -File | Select-Object -Expand FullName
    }
    else {
        $Links = Get-ChildItem $Path -Force -Filter *.lnk -File -Recurse | Select-Object -Expand FullName
    }

    Get-LinkContent $Links
}
