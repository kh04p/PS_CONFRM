$HistoryPaths = Get-Content -path D:\PowerBI_ConferenceRoom\SourceData\HistoryPaths.txt                               #These file paths will not contain .csv at the end for export reasons.
$DevicePaths = Get-Content -path D:\PowerBI_ConferenceRoom\SourceData\DevicePaths.txt 

foreach($path in $HistoryPaths) {
    $inputpath = $path + '.csv'                                                                #Adds .csv to end of file path before importing.                                                                  

    $inputcsv = Import-Csv $inputpath | Sort-Object "Computer Name","Offline Dates" -Unique    #Sort object and remove row if same value for both "Computer Name" and "Offline Dates".
                                                                                               #Might need to sort again (without -Unique) to return document back to original sorting order.
    $inputcsv | Format-Table                                                                   #Writes to terminal in table form.

    $outputpath = $path + '_sorted.csv'                                                        #Appends to file name before exporting.

    $inputcsv | Export-Csv -Path $outputpath -NoTypeInformation -Force                         #Exports and overwrites any existing file with same name.

    Write-Host 'Removed duplicates for'$inputpath
}

foreach($path in $DevicePaths) {
    $inputpath = $path + '.csv'

    $inputcsv = Import-Csv $inputpath | Sort-Object "Computer","Status","Updated" -Unique

    $inputcsv | Format-Table

    $outputpath = $path + '_sorted.csv'

    $inputcsv | Export-Csv -Path $outputpath -NoTypeInformation -Force

    Write-Host 'Removed duplicates for'$inputpath
}

#Read-Host -Prompt 'Press Enter to exit'