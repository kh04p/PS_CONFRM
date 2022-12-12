Clear-Content D:\PowerBI_ConferenceRoom\SourceData\ConfDevices.txt -Force
$conf = Get-ADComputer -Filter * -SearchBase "OU=Conference_Rooms,OU=Clients,OU=Workstations,OU=Windows_10_Production,OU=Clients,OU=Workstations,DC=corp,DC=costco,DC=com"
$conf.Name | Sort-Object | Out-file D:\PowerBI_ConferenceRoom\SourceData\ConfDevices.txt -Force
$byod = "BYOD" | Out-File D:\PowerBI_ConferenceRoom\SourceData\ConfDevices.txt -Append -Force
$byodonly = "BYOD Only" | D:\PowerBI_ConferenceRoom\SourceData\ConfDevices.txt -Append -Force