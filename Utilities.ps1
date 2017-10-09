function CheckFilesAgainstCsv {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,HelpMessage='the path to the csv')]
        [ValidateNotNullOrEmpty()]
        [String]$CSVPath,

        [Parameter(Mandatory,HelpMessage='The column names which contain the elements of the filename')]
        [ValidateNotNullOrEmpty()]
        [String[]]$ColumnsWithFileNameParts,

        [Parameter(HelpMessage='The column names which contain the elements of the filename')]
        [ValidateNotNullOrEmpty()]
        [String]$Extension
    )

    if ($Extension -and $Extension.Substring(0,1) -ne ".") {$Extension = ".$Extension"}

    $CSV = Import-Csv -Path $CSVPath -Encoding Unicode

    $FileNames = $CSV.GetEnumerator() | ForEach-Object {
        $Map = $_
        ($ColumnsWithFileNameParts | % {$Map."$($_)"}) -join "\"
    }

    [String[]]$Found = $null
    [String[]]$Directories = $null
    [String[]]$Failed = $null

    $FileNames | % {
        if (Test-Path -Path $_) {
            if ($_.Substring($_.Length-1) -eq "\") {
                $Directories += $_
            } else {   
                $Found += $_
            }
        } elseif (Test-Path -Path "$($_)$Extension") {
            $Found += "$($_)$Extension"
        } else {
            $Failed += "$($_)$Extension"
        }
    }

    $Directories | % {Write-Warning "$_ is a directory"}
    $Found | % {Write-Host $_ -BackgroundColor Black -ForegroundColor Cyan}
    $Failed | % {Write-Host $_ -BackgroundColor Black -ForegroundColor Red}

}