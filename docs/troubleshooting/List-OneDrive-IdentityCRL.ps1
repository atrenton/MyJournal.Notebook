# List-OneDrive-IdentityCRL.ps1

$account = @{
    Name = 'MicrosoftAccount';
    Expression = {$_}
}
$cid = @{
    Name = 'CID';
    Expression = { $item.OpenSubKey($_).GetValue("cid") }
}
$webcredtype = @{
    Name = 'WebCredType';
    Expression = { $item.OpenSubKey($_).GetValue("webcredtype") }
}

$item = Get-Item "HKCU:\Software\Microsoft\IdentityCRL\UserExtendedProperties"
$item.GetSubKeyNames() | Select-Object $account, $cid, $webcredtype |
    Out-GridView -Title 'OneDrive Identity Cross Reference List' -Wait
