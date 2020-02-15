# OneNote-Library.ps1
#Requires -Version 5.1

Function not-exist { -not (Test-Path $args) }
Set-Alias !exist not-exist
Set-Alias exist Test-Path

Function Display-OneNote-Version
{
    param(
        [ValidateScript({
            if( $_ -match 'Office\d\d' )
            {
                $true
            }
            else
            {
                throw "Invalid value: $_"
            }
        })][string] $OfficeProductName
    )
    [hashtable]$version_lookup = @{
        12 = '2007'
        14 = '2010'
        15 = '2013'
        16 = '2016'
    }

    $len = $OfficeProductName.Length
    [int]$i = $OfficeProductName.Substring($len - 2)
    echo "`r`nOneNote $($version_lookup[$i]) is installed."
}

Function Get-OneNote-Bitness
{
    param (
        [Parameter(Mandatory=$true)][string]$ExeFilePath
    )
    if (exist $ExeFilePath) {
        # The following code was inspired by:
        # https://github.com/guyrleech/Microsoft/blob/master/Get%20file%20bitness.ps1
        [int]$MACHINE_OFFSET = 4
        [int]$PE_POINTER_OFFSET = 60

        [hashtable]$bitness = @{
            0x014c = '32-bit'
            0x8664 = '64-bit'
        }
        $data = New-Object System.Byte[] 4096
        $stream = New-Object System.IO.FileStream -ArgumentList $ExeFilePath,Open,Read
        $stream.Read($data, 0, $data.Count) | Out-Null
        $stream.Close()

        [int]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
        [int]$typeOffset = $PE_HEADER_ADDR + $MACHINE_OFFSET
        [uint16]$machineType = [System.BitConverter]::ToUInt16($data, $typeOffset)

        return $bitness[[int]$machineType]
    } else {
        throw 'ONENOTE.EXE not found'
    }
}

Function Press-Any-Key
{
    Write-Host 'Press any key to continue. . .' -NoNewline
    $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown') | Out-Null
}
