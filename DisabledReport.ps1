function Save-CSVasExcel {
    param (
        [string]$CSVFile = $(Throw 'No file provided.')
    )
    
    BEGIN {
        function Resolve-FullPath ([string]$Path) {    
            if ( -not ([System.IO.Path]::IsPathRooted($Path)) ) {
                # $Path = Join-Path (Get-Location) $Path
                $Path = "$PWD\$Path"
            }
            [IO.Path]::GetFullPath($Path)
        }

        function Release-Ref ($ref) {
            ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        
        $CSVFile = Resolve-FullPath $CSVFile
        $xl = New-Object -ComObject Excel.Application
    }

    PROCESS {
        $wb = $xl.workbooks.open($CSVFile)
        $xlOut = $CSVFile -replace '\.csv$', '.xlsx'
        
        # can comment out this part if you don't care to have the columns autosized
        $ws = $wb.Worksheets.Item(1)
        $range = $ws.UsedRange 
        [void]$range.EntireColumn.Autofit()

        $num = 1
        $dir = Split-Path $xlOut
        $base = $(Split-Path $xlOut -Leaf) -replace '\.xlsx$'
        $nextname = $xlOut
        while (Test-Path $nextname) {
            $nextname = Join-Path $dir $($base + "-$num" + '.xlsx')
            $num++
        }

        $wb.SaveAs($nextname, 51)
    }

    END {
        $xl.Quit()
    
        $null = $ws, $wb, $xl | % {Release-Ref $_}

        # del $CSVFile
    }
}
$currentDate = Get-Date
$currentDate = $currentDate.ToString('MM-dd-yyyy_hh-mm-ss')
$file = "C:\scriptReports\DisabledPCs_$currentDate" + ".csv"
Get-ADComputer -Filter * -SearchBase "OU=Disabled_Computers,OU=Computers,OU=,DC=something,DC=dot,DC=com" -Properties Description | select Name,Description |Export-CSV -Path $file -NoTypeInformation
Save-CSVasExcel $file
$newfile = $file -replace '\.csv$', '.xlsx'
Remove-Item $file
Send-MailMessage -To “coolbreeze822@email.com" -From “NO-REPLY@email.com" -SMTPServer mail.email.com -Subject “Disabled Computers” -Body “Please review this weeks report of Computers in the Disabled_Computer OU and do what you can to help clean it out. `n `n Please Do Not Respond To This Email `n `n Thank you ” -Attachments $newfile -Priority High


