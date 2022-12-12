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
Write-Host '░░░░░░░░░░░░AD Computer Check░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black
Write-Host '░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░' -BackgroundColor White -ForegroundColor Black



foreach($comp in $computers) {
    $exists = $true
    try {$ADComputer = Get-ADComputer $comp -Server 'corp.costco.com'}
    catch {
        try {$ADComputer = Get-ADComputer $comp -Server 'systems.costco.com'}
        catch {
            Write-Host $comp 'is ' -NoNewline
            Write-Host 'NOT in Active Directory.' -ForegroundColor Red -BackgroundColor Black
            $exists = $false
        }

        if ($exists) {
            $ping = Test-Connection $comp -count 1 -quiet 
            Write-Host $comp -NoNewLine 
            Write-Host ' is '-NoNewLine
    
            if($ping){
                Write-Host 'online.' -Foregroundcolor Green -BackgroundColor Black -NoNewline
                Write-Host ' in ' -NoNewline
                Write-Host 'systems.costco.com' -ForegroundColor Green

            }else{
                Write-Host 'offline.' -Foregroundcolor Red -BackgroundColor Black -NoNewline
                Write-Host ' in ' -NoNewline
                Write-Host 'systems.costco.com' -ForegroundColor Green 
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

        }else{
            Write-Host 'offline' -Foregroundcolor Red -BackgroundColor Black -NoNewline
            Write-Host ' in ' -NoNewline
            Write-Host 'corp.costco.com' -ForegroundColor Blue   
        }
    }
}