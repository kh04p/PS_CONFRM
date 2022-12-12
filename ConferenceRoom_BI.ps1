Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░[Conference Room Health Check]░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black

#1. GETS SOURCE LIST OF ALL DEVICES
$Computers = Get-Content -Path D:\PowerBI_ConferenceRoom\SourceData\ConfDevices.txt 

#2. GET TIME WHEN SCRIPT STARTS
$date = Get-Date -Format "MM/dd/yyyy HH:mm:ss"  
$dateShort = Get-Date -Format "MM/dd/yyyy"

#3. CHECK IF MONTH HAS ENDED AND CREATE NEW CSV IF NEW MONTH
$timeCSV = Import-Csv 'D:\PowerBI_ConferenceRoom\Results\OfflineDevices.csv' | Select-Object -Last 1
$timeCSV_month = (Get-Date -Date $timeCSV.Updated).Month

$timeCheck = (Get-Date).Date
if ($timeCheck.Month -ne $timeCSV_month) {
    Write-Host $timeCSV_month
    Write-Host 'New month -> Renaming history file...'
    $NewName = 'OfflineHistory_' + $timeCheck.Month + $timeCheck.Year + '.csv'
    Rename-Item -Path D:\PowerBI_ConferenceRoom\Results\OfflineHistory.csv -NewName $NewName
    Write-Host 'New month -> Renaming old file...'
} else {
    Write-Host 'Same month -> Continuing...'
}

clear-content D:\PowerBI_ConferenceRoom\Results\OfflineDevices.csv -Force
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
                $ADComputer=[pscustomobject]@{
                    'ComputerName'=$Computer
                    'Status'='FALSE'
                }

                #export list of computers not in AD
                $ADComputer | Export-CSV D:\PowerBI_ConferenceRoom\Results\AD.csv -Append -NoTypeInformation -Force

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
                $ComputerOffline | Export-CSV D:\PowerBI_ConferenceRoom\Results\OfflineHistory.csv -Append -NoTypeInformation -Force
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
            $ComputerOffline | Export-CSV D:\PowerBI_ConferenceRoom\Results\OfflineHistory.csv -Append -NoTypeInformation -Force  
        }
    }
     
    #B. CREATES CUSTOM OBJECT FOR EXPORTING  
    $Computer =[pscustomobject]@{                                                  #creates custom object once status has been retrieved to export to CSV
        'Computer' = $Computer                                                     #each field will be a header with its own column
        'Status' = $Status 
        'Updated' = $date    
    }

    #C. EXPORTS INTO LIST OF DEVICES AND STATUSES
    $Computer | Export-CSV D:\PowerBI_ConferenceRoom\Results\OfflineDevices.csv -Append -NoTypeInformation -Force  
   
}

#5. RUNS SCRIPT FOR REGIONAL ROOMS AND CLEANS UP RESULTS
$PSScriptRoot
& $PSScriptRoot\RegionalConferenceRoom_BI.ps1                                      #automatically runs the one for Regional rooms in the same directory
& $PSScriptRoot\RemoveDuplicates.ps1                                               #removes duplicates in result files

#6. PAUSES SCRIPT FOR USER TO READ RESULTS IF NEEDED
Read-Host -Prompt 'Press Enter to exit'