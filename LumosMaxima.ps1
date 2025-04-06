$taskNameMorning = "AutoSwitchAppThemeMorning"
$taskNameEvening = "AutoSwitchAppThemeEvening"
$taskNameLogon = "AutoSwitchAppThemeOnLogon"
$scriptPath = $MyInvocation.MyCommand.Definition

$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($CurrentIdentity)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    $arguments = $MyInvocation.MyCommand.Definition
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$arguments`"" -Verb RunAs
    exit
}

$taskExistsLogon = Get-ScheduledTask -TaskName $taskNameLogon -ErrorAction SilentlyContinue
if (-not $taskExistsLogon) {
    schtasks /Create `
        /TN $taskNameLogon `
        /TR "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" `
        /SC ONLOGON `
        /RL HIGHEST `
        /F
}

$taskExistsMorning = Get-ScheduledTask -TaskName $taskNameMorning -ErrorAction SilentlyContinue
if (-not $taskExistsMorning) {
    schtasks /Create `
        /TN $taskNameMorning `
        /TR "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" `
        /SC DAILY `
        /ST 06:00 `
        /RL HIGHEST `
        /F
}

$taskExistsEvening = Get-ScheduledTask -TaskName $taskNameEvening -ErrorAction SilentlyContinue
if (-not $taskExistsEvening) {
    schtasks /Create `
        /TN $taskNameEvening `
        /TR "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" `
        /SC DAILY `
        /ST 18:00 `
        /RL HIGHEST `
        /F

    Write-Host "Your apps that use default colour themes will now automatically switch themes at 6 AM and 6 PM - keeping your eyes nice and fresh!"
}

$hour = (Get-Date).Hour

if ($hour -ge 6 -and $hour -lt 18) {
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -Value 1
}
else {
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -Value 0
}