if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams -Force
}

Connect-MicrosoftTeams

$Excel = New-Object -ComObject Excel.Application

Get-ChildItem -Filter "*.xls" | Foreach-Object{
    Write-Host $_.FullName -ForegroundColor Yellow
    $Excel.Workbooks.Open($_.FullName).Sheets | ForEach-Object{
        $TeamName = $_.Range("W2").Value2
    
        $NewTeam = Get-Team -DisplayName $TeamName
        if (-Not $NewTeam) {
            $NewTeam = New-Team -DisplayName $TeamName -Template "EDU_Class" -ErrorAction Stop
            Write-Host "New team created: $TeamName"
        }
        $Students = $_.Range("B:B").Value2 | Where {$_ -Match "^\d+$"} | ForEach-Object {
                $StudentId = $_

                try {
                    Add-TeamUser -GroupId $NewTeam.GroupId -User "$StudentId@alanyauniversity.edu.tr" -Role Member
                    Write-Host "$StudentId@alanyauniversity.edu.tr -> $($NewTeam.DisplayName)"  -ForegroundColor Green
                }
                catch {
                    Write-Host "$StudentId@alanyauniversity.edu.tr : $_" -ForegroundColor Red
                }
        }
        Write-Host "$($NewTeam.DisplayName) processed successfully!" -ForegroundColor Yellow
    }
}

Write-Host "All files processed successfully!"
Disconnect-MicrosoftTeams
Pause
