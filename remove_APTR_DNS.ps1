$ListOfHosts = @("Server1","Server2","Server3")
foreach ($HostToDelete in $ListOfHosts){
$DNSDirectZone = $env:userdnsdomain
$DNSServer = $env:logonserver -replace '\\',''

$DNSARecord = Resolve-DnsName $HostToDelete
$AHostName = $DNSARecord.Name -replace $DNSDirectZone,"" -replace "\.$",""

$DNSPtrRecord = Resolve-DnsName $DNSARecord.IPAddress
$DNSReverseZone = (Get-DnsServerZone -ComputerName $DNSServer | ?{$DNSPtrRecord.Name -match $_.ZoneName -and $_.IsDsIntegrated -eq $true}).ZoneName

$PtrHostName = $DNSARecord.IPAddress -split "\."
[array]::Reverse($PtrHostName)
$PtrHostName = $PtrHostName -join "." -replace $DNSReverseZoneSuffix,"" -replace "\.$",""

if ($DNSPtrRecord.NameHost -eq $DNSARecord.Name) {
Remove-DnsServerResourceRecord -ComputerName $DNSServer -ZoneName $DNSReverseZone -Name $PtrHostName -RRType Ptr -Confirm:$false -Force
}

Remove-DnsServerResourceRecord -ComputerName $DNSServer -ZoneName $DNSDirectZone -Name $AHostName -RRType A -Confirm:$false -Force
}