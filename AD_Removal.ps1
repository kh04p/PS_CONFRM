$computers = Get-Content -Path 'C:\PowerBI\PowerBI_ConferenceRoom\SourceData\MergedDevices.txt'

Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░▄▄▀▀▀▀▀▀▀▀▀▀▄▄█▄░░░░▄░░░░█░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░█▀░░░░░░░░░░░░░▀▀█▄░░░▀░░░░░░░░░▄░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░▄▀░░░░░░░░░░░░░░░░░▀██░░░▄▀▀▀▄▄░░▀░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░▄█▀▄█▀▀▀▀▄░░░░░░▄▀▀█▄░▀█▄░░█▄░░░▀█░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░▄█░▄▀░░▄▄▄░█░░░▄▀▄█▄░▀█░░█▄░░▀█░░░░█░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '▄█░░█░░░▀▀▀░█░░▄█░▀▀▀░░█░░░█▄░░█░░░░█░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '██░░░▀▄░░░▄█▀░░░▀▄▄▄▄▄█▀░░░▀█░░█▄░░░█░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '██░░░░░▀▀▀░░░░░░░░░░░░░░░░░░█░▄█░░░░█░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '██░░░░░░░░░░░░░░░░░░░░░█░░░░██▀░░░░█▄░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '██░░░░░░░░░░░░░░░░░░░░░█░░░░█░░░░░░░▀▀█▄' -BackgroundColor White -ForegroundColor Black
Write-Host '██░░░░░░░░░░░░░░░░░░░░█░░░░░█░░░░░░░▄▄██' -BackgroundColor White -ForegroundColor Black
Write-Host '░██░░░░░░░░░░░░░░░░░░▄▀░░░░░█░░░░░░░▀▀█▄' -BackgroundColor White -ForegroundColor Black
Write-Host '░▀█░░░░░░█░░░░░░░░░▄█▀░░░░░░█░░░░░░░▄▄██' -BackgroundColor White -ForegroundColor Black
Write-Host '░▄██▄░░░░░▀▀▀▄▄▄▄▀▀░░░░░░░░░█░░░░░░░▀▀█▄' -BackgroundColor White -ForegroundColor Black
Write-Host '░░▀▀▀▀░░░░░░░░░░░░░░░░░░░░░░█▄▄▄▄▄▄▄▄▄██' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black


Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░ Conference Room Check ░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black

Clear-Content D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt -Force
Clear-content D:\PowerBI_ConferenceRoom\Results\AD_Add.txt -Force

#CHECKS IF ORIGINAL COMPUTER NAME IS PINGABLE AND IN AD

#IF NOT PINGABLE OR IN AD:
#CHANGES THIN CLIENT NAME'S FIRST LETTER FROM 'T' TO 'W' AND TEST
#CHANGES PC NAME TO HAVE AN 'A' AT THE END AND TEST

:outer1 foreach($comp in $computers) {
    
    if ($comp.ToLower() -eq "byod" -or $comp.ToLower() -eq "byod only") {
        continue outer1
    }

    $exists = $true
    try {$ADComputer = Get-ADComputer $comp -Server 'corp.costco.com'}
    catch {
        try {$ADComputer = Get-ADComputer $comp -Server 'systems.costco.com'}
        catch {
            Write-Host $comp 'is ' -NoNewline
            Write-Host 'NOT in Active Directory.' -ForegroundColor Red -BackgroundColor Black
            $exists = $false

            $firstChar = $comp[0]
            if ($firstChar -eq 'T' -or $firstChar -eq 'T') {
                $str = 'w' + $comp.Substring(1)
                $ping = Test-Connection $str -count 1 -quiet
                if ($ping) {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                    Write-Host ' Will need to add to device list. '
                    Write-Host
                    $str | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Add.txt -Append -Force
                }
                else {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                    Write-Host
                }
            }

            elseif ($firstChar -eq 'W' -or $firstChar -eq 'w') {
                $str = $comp + 'a'
                $ping = Test-Connection $str -count 1 -quiet
                if ($ping) {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                    Write-Host ' Will need to add to device list. '
                    Write-Host
                    $str | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Add.txt -Append -Force
                }
                else {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                    Write-Host
                }
            }
        }

        if ($exists) {
            $ping = Test-Connection $comp -count 1 -quiet 
            Write-Host $comp -NoNewLine 
            Write-Host ' is '-NoNewLine
    
            if($ping){
                Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                Write-Host ' in ' -NoNewline
                Write-Host 'systems.costco.com' -ForegroundColor Green
                Write-Host

            }else{
                Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black -NoNewline
                Write-Host ' in ' -NoNewline
                Write-Host 'systems.costco.com' -ForegroundColor Green 

                $firstChar = $comp[0]
                if ($firstChar -eq 't' -or $firstChar -eq 'T') {
                    $str = 'w' + $comp.substring(1)

                    $ping = Test-Connection $str -count 1 -quiet
                    if ($ping) {
                        Write-Host $str 'is ' -NoNewLine
                        Write-Host 'online. ' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                        Write-Host 'Will need to delete ' -NoNewLine
                        Write-Host $comp -Foregroundcolor White -BackgroundColor Black
                        Write-Host "`r`n"
                        $comp | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt -Append -Force
                    }
                    else {
                        Write-Host $str 'is also ' -NoNewLine
                        Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                        Write-Host
                    }
                }
                else {
                    $str = $comp + 'a'
                    $ping = Test-Connection $str -count 1 -quiet
                    if ($ping) {
                        Write-Host $str 'is ' -NoNewLine
                        Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                        Write-Host ' Will need to delete ' -NoNewLine
                        Write-Host $comp -Foregroundcolor White -BackgroundColor Black
                        Write-Host "`r`n"
                        $comp | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt -Append -Force
                    }
                    else {
                        Write-Host $str 'is also ' -NoNewLine
                        Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                        Write-Host
                    }
                }
            }
        }
    }

    if ($exists) {       
        $ping = Test-Connection $comp -count 1 -quiet 
        Write-Host $comp -NoNewLine 
        Write-Host ' is '-NoNewLine
    
        if($ping){
            Write-Host 'online' -Foregroundcolor Green -BackgroundColor Black -NoNewline
            Write-Host ' in ' -NoNewline
            Write-Host 'corp.costco.com' -ForegroundColor Blue 
            Write-Host  

        }else{
            Write-Host 'offline' -Foregroundcolor Red -BackgroundColor Black -NoNewline
            Write-Host ' in all domains.' 
             
            $firstChar = $comp[0]
            if ($firstChar -eq 't' -or $firstChar -eq 'T') {
                $str = 'w' + $comp.substring(1)

                $ping = Test-Connection $str -count 1 -quiet
                if ($ping) {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                    Write-Host ' Will need to delete ' -NoNewLine
                    Write-Host $comp -Foregroundcolor White -BackgroundColor Black
                    Write-Host "`r`n"
                    $comp | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt -Append -Force
                }
                else {
                    Write-Host $str 'is also ' -NoNewLine
                    Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                    Write-Host
                }
            }
            else {
                $str = $comp + 'a'
                $ping = Test-Connection $str -count 1 -quiet
                if ($ping) {
                    Write-Host $str 'is ' -NoNewLine
                    Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                    Write-Host ' Will need to delete ' -NoNewLine
                    Write-Host $comp -Foregroundcolor White -BackgroundColor Black -NoNewline
                    Write-Host "`r`n"
                    $comp | Out-File -FilePath D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt -Append -Force
                }
                else {
                    Write-Host $str 'is also ' -NoNewLine
                    Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black
                    Write-Host
                }
            }
        }
    }
}

#FILTER LIST WITH ROOMS COMPLETED IN CONFERENCE ROOM PROJECT TRACKER SPREADSHEET: https://drive.google.com/u/0/open?id=1L7dOVDT2nEU7oqNp1EXFagvSuiyR93byWAXB_vVrjzk

$completedRooms = Get-Content D:\PowerBI_ConferenceRoom\SourceData\Completed_Rooms.txt
$toBeRemoved = Get-Content D:\PowerBI_ConferenceRoom\Results\AD_Delete.txt

$number = 0

:outer2 foreach ($computer1 in $toBeRemoved) {
    $number = $number + 1

    #CHECK IF THIN CLIENT:
    $firstChar = $computer1[0]
    if ($firstChar -eq 't' -or $firstChar -eq 'T') {
        $oldComputer1 = $computer1
        $computer1 = 'w' + $computer1.Substring(1)

        foreach ($computer2 in $completedRooms) {
            if ($computer1 -eq $computer2) {
                Write-Host '[' $number ']' $oldComputer1 'is ready to be deleted.' -ForegroundColor Green -BackgroundColor Black
                $oldComputer1 | Out-File D:\PowerBI_ConferenceRoom\AD_Removal\AD_Delete_Filtered.txt -Append -Force
                continue outer2
            }
        }

        Write-Host '[' $number ']' $oldComputer1 'is NOT ready to be deleted!' -ForegroundColor Red -BackgroundColor Black
    }

    #CHECK IF PC:
    else {
        foreach ($computer2 in $completedRooms) {
            if ($computer1 -eq $computer2) {
                Write-Host '[' $number ']' $oldComputer1 'is ready to be deleted.' -ForegroundColor Green -BackgroundColor Black
                $oldComputer1 | Out-File D:\PowerBI_ConferenceRoom\AD_Removal\AD_Delete_Filtered.txt -Append -Force
                continue outer2
            }            
        }

        $oldComputer1 = $computer1
        $computer1 = $oldComputer1 + 'a'

        foreach ($computer2 in $completedRooms) {
            if ($computer1 -eq $computer2) {
                Write-Host '[' $number ']' $oldComputer1 'is ready to be deleted.' -ForegroundColor Green -BackgroundColor Black
                $oldComputer1 | Out-File D:\PowerBI_ConferenceRoom\AD_Removal\AD_Delete_Filtered.txt -Append -Force
                continue outer2
            }
            
        }

        Write-Host $number ':' $oldComputer1 'is NOT ready to be deleted!' -ForegroundColor Red -BackgroundColor Black

    }
    
}

#FINAL LIST READY TO BE DELETED FROM AD

$finalList = Get-Content D:\PowerBI_ConferenceRoom\AD_Removal\AD_Delete_Filtered.txt

Write-Host
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░ FINAL DEVICE LIST ░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host
Write-Host 'Do you want to remove the following devices from Active Directory?'
Write-Host
$counter = 0

foreach ($computer in $finalList) {
    $counter = $counter + 1
    Write-Host '[' $counter ']' $computer
}

Write-Host
$loop = $true

while ($loop) {
    $user = (Read-Host -Prompt 'Yes (Y) or No (N)').ToLower()

    if ($user -eq 'y' -or $user -eq 'yes') {
        $counter = 0
        foreach ($finalComputer in $finalList) {
            $counter = $counter + 1
            Get-ADComputer $finalComputer | Remove-ADObject -Recursive -Confirm:$false
            Write-Host 'Removed' $finalComputer '(' $counter ')'
        }

        Write-Host
        Write-Host 'All' $counter 'devices above have been deleted from Active Directory.'
        $loop = $false
    }

    elseif ($user -eq 'n' -or $user -eq 'no') {
        Write-Host 'No devices have been deleted, ending script.'
        $loop = $false
    }

    else {
        Write-Host 'Try again, you got one job man.'
        Write-Host
        Write-Host 'Do you want to remove the following devices from Active Directory?'
    }
}

foreach ($computer in $finalList) {
    try {$ADComputer = Get-ADComputer $computer -Server 'corp.costco.com'}
    catch {
        Write-Host $computer 'is deleted.'
    }
}

#RENAME FILE TO BACK UP
$timeCheck = (Get-Date).Date
$NewName = 'AD_Deleted_Filtered' + $timeCheck.Month + $timeCheck.Year + '.txt'
try {
    Rename-Item -Path D:\PowerBI_ConferenceRoom\AD_Removal\AD_Delete_Filtered.txt -NewName $NewName
} catch {
    Write-Host 'Unable to rename AD_Deleted_Filtered file.'
}

Write-Host
Write-Host 'Renamed file successfully to' + $NewName
Write-Host 'Adios.'