#1. GETS SOURCE LIST OF ALL DEVICES
$Computers = Get-Content -Path D:\PowerBI_ConferenceRoom\SourceData\RegionalConfDevices.txt 

#2. GET TIME WHEN SCRIPT STARTS
$date = Get-Date -Format "MM/dd/yyyy HH:mm:ss"  
$dateShort = Get-Date -Format "MM/dd/yyyy"

#3. CHECK IF MONTH HAS ENDED AND CREATE NEW CSV IF NEW MONTH
$timeCSV = Import-Csv 'D:\PowerBI_ConferenceRoom\Results\RegionalOfflineDevices.csv' | Select-Object -Last 1
Write-Host $timeCSV
$timeCSV_month = (Get-Date -Date $timeCSV.Updated).Month

$timeCheck = (Get-Date).Date
if ($timeCheck.Month -ne $timeCSV_month) {
    Write-Host $timeCSV_month
    Write-Host 'New month -> Renaming history file...'
    $NewName = 'OfflineHistory_' + $timeCheck.Month + $timeCheck.Year + '.csv'
    Rename-Item -Path D:\PowerBI_ConferenceRoom\Results\RegionalOfflineHistory.csv -NewName $NewName
} else {
    Write-Host 'Same month -> Continuing...'
}

clear-content D:\PowerBI_ConferenceRoom\Results\RegionalOfflineDevices.csv -Force
clear-content D:\PowerBI_ConferenceRoom\Results\AD.csv

#4. MAIN SCRIPT
foreach($Computer in $Computers){
    

    #A. PINGS EACH COMPUTER
    $ping = Test-Connection $Computer -count 1 -quiet                      

    if ($ping -or $Computer -eq "BYOD" -or $Computer -eq "BYOD only") {            #auto flags BYOD rooms as online
        $Status = 'online'                                                         #assigns online status to variable
        Write-Host $Computer 'is ' -NoNewline
        Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black 
    }else{

        #A1. IF COMPUTER IS NOT ONLINE, CHECK ACTIVE DIRECTORY
        $exists = $true
        try {$AComputer = Get-ADComputer $Computer -Server 'corp.costco.com'} catch {
            try {$ADComputer = Get-ADComputer $Computer -Server 'systems.costco.com'} catch {
                $exists = $false
                $ADComputer =[pscustomobject]@{
                    'ComputerName'=$Computer
                    'Status'='FALSE'
                }

                #export list of computers not in AD
                $ADComputer | Export-CSV C:\PowerBI\PowerBI_ConferenceRoom\Results\AD.csv -Append -NoTypeInformation -Force

                $Status = 'Not_AD'                                                 #status if computer cannot be found in AD
                Write-Host $Computer 'is ' -NoNewline
                Write-Host 'NOT in AD' -Foregroundcolor Blue -BackgroundColor Black -NoNewline
                Write-Host ' on ' -NoNewline
                Write-Host $date -Foregroundcolor Red -BackgroundColor Black                                
            }

            if ($exists) {
                $Status = 'offline'                                                #status if computer can be found in AD
                Write-Host $Computer 'is ' -NoNewline
                Write-Host 'offline' -Foregroundcolor Red -BackgroundColor Black -NoNewline
                Write-Host ' on ' -NoNewline
                Write-Host $date -Foregroundcolor Red -BackgroundColor Black

                $ComputerOffline =[pscustomobject]@{                               #custom object to export which devices went offline for logging purposes
                    'Computer Name' = $Computer                                    #each field here will automatically be a different header in exported CSV file
                    'Offline Dates' = $dateShort
                }
                $ComputerOffline | Export-CSV D:\PowerBI_ConferenceRoom\Results\RegionalOfflineHistory.csv -Append -NoTypeInformation -Force
            }
        }

        if ($exists) {
            $Status = 'offline'                                                    #status if computer can be found in AD
            Write-Host $Computer 'is ' -NoNewline
            Write-Host 'offline' -Foregroundcolor Red -BackgroundColor Black -NoNewline
            Write-Host ' on ' -NoNewline
            Write-Host $date -Foregroundcolor Red -BackgroundColor Black

            $ComputerOffline =[pscustomobject]@{                                   #custom object to export which devices went offline for logging purposes
                'Computer Name' = $Computer                                        #each field here will automatically be a different header in exported CSV file
                'Offline Dates' = $dateShort
            }
            $ComputerOffline | Export-CSV D:\PowerBI_ConferenceRoom\Results\RegionalOfflineHistory.csv -Append -NoTypeInformation -Force  
        }
    }
     
    #B. CREATES CUSTOM OBJECT FOR EXPORTING  
    $Computer =[pscustomobject]@{                                                  #creates custom object once status has been retrieved to export to CSV
        'Computer' = $Computer                                                     #each field will be a header with its own column
        'Status' = $Status 
        'Updated' = $date    
    }

    #C. EXPORTS INTO LIST OF DEVICES AND STATUSES
    $Computer | Export-CSV D:\PowerBI_ConferenceRoom\Results\RegionalOfflineDevices.csv -Append -NoTypeInformation -Force  
   
}