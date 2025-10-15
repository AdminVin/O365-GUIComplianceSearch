<#
[>] Change Log
2025-10-15 - v1.1
    - Added "-EnableSearchOnlySession" to resolve purge errors.
2025-04-27 - v1.0
    - Initial Release.
#>

<# Prerequ#>
# Exchange Online
if (!(Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Install-Module ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber
}

# GUI Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "O365 - GUI Compliance Search (AdminVin)"
$form.Size = New-Object System.Drawing.Size(550, 550)
$form.StartPosition = "CenterScreen"

$labels = @(
    "Search Name",
    "Sender Email",
    "Sender Email Notes",
    "Scope (subject/body)",
    "Search Term",
    "Search Term Note",
    "Start Date",
    "End Date",
    "Purge Type",
    "Delete Search",
    "Delete Search Note"
)

$controls = @{}
$buttons = @{}
$y = 10

foreach ($label in $labels) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $label
    $lbl.Location = New-Object System.Drawing.Point(10, $y)
    $lbl.Size = New-Object System.Drawing.Size(140, 20)
    $form.Controls.Add($lbl)

    if ($label -eq "Scope (subject/body)") {
        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Location = New-Object System.Drawing.Point(150, $y)
        $dropdown.Size = New-Object System.Drawing.Size(150, 20)
        $dropdown.Items.AddRange(@("subject", "body"))
        $form.Controls.Add($dropdown)
        $controls[$label] = $dropdown
        $y += 40
    }
    elseif ($label -eq "Sender Email Notes") {
        $note = New-Object System.Windows.Forms.Label
        $note.Text = "Use * for any sender (OR) Wildcard: vincent*"
        $note.Location = New-Object System.Drawing.Point(148, $y)
        $note.Size = New-Object System.Drawing.Size(350, 20)
        $form.Controls.Add($note)
        $y += 30
    }
    elseif ($label -eq "Search Term Note") {
        $note = New-Object System.Windows.Forms.Label
        $note.Text = "Use 'subject' and  * for searching all messages."
        $note.Location = New-Object System.Drawing.Point(148, $y)
        $note.Size = New-Object System.Drawing.Size(350, 20)
        $form.Controls.Add($note)
        $y += 30
    }
    elseif ($label -eq "Delete Search Note") {
        $note = New-Object System.Windows.Forms.Label
        $note.Text = "Checking 'yes' will delete the content search from the compliance portal after purging."
        $note.Location = New-Object System.Drawing.Point(148, $y)
        $note.Size = New-Object System.Drawing.Size(350, 40)
        $form.Controls.Add($note)
        $y += 50
    }
    elseif ($label -eq "Purge Type") {
        $chkSoft = New-Object System.Windows.Forms.CheckBox
        $chkSoft.Text = "SoftDelete"
        $chkSoft.Location = New-Object System.Drawing.Point(155, $y)
        $form.Controls.Add($chkSoft)
        $controls["SoftDelete"] = $chkSoft

        $chkHard = New-Object System.Windows.Forms.CheckBox
        $chkHard.Text = "HardDelete"
        $chkHard.Location = New-Object System.Drawing.Point(265, $y)
        $form.Controls.Add($chkHard)
        $controls["HardDelete"] = $chkHard

        $y += 40
    }
    elseif ($label -eq "Delete Search") {
        $chkDelete = New-Object System.Windows.Forms.CheckBox
        $chkDelete.Text = "Yes"
        $chkDelete.Location = New-Object System.Drawing.Point(155, $y)
        $form.Controls.Add($chkDelete)
        $controls["Delete Search"] = $chkDelete
        $y += 40
    }
    elseif ($label -in @("Start Date", "End Date")) {
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(150, $y)
        $txt.Size = New-Object System.Drawing.Size(140, 20)
        $txt.ReadOnly = $true
        $form.Controls.Add($txt)
        $controls[$label] = $txt

        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = "Select"
        $btn.Size = New-Object System.Drawing.Size(60, 20)
        $btn.Location = New-Object System.Drawing.Point(300, $y)
        $form.Controls.Add($btn)
        $buttons[$label] = $btn

        # Optional Label
        $optionalLbl = New-Object System.Windows.Forms.Label
        $optionalLbl.Text = "(Optional)"
        $optionalLbl.Location = New-Object System.Drawing.Point(375, $y)
        $optionalLbl.Size = New-Object System.Drawing.Size(80, 20)
        $form.Controls.Add($optionalLbl)

        $y += 40
    }
    else {
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(150, $y)
        $txt.Size = New-Object System.Drawing.Size(350, 20)
        $form.Controls.Add($txt)
        $controls[$label] = $txt
        $y += 40
    }
}

# Calendar button handling
$buttons["Start Date"].Add_Click({
    $calendarForm = New-Object System.Windows.Forms.Form
    $calendarForm.Text = "Select Start Date"
    $calendarForm.Size = New-Object System.Drawing.Size(250, 280)
    $calendarForm.StartPosition = "CenterScreen"

    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.MaxSelectionCount = 1
    $calendar.Dock = "Top"
    $calendarForm.Controls.Add($calendar)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Location = New-Object System.Drawing.Point(80, 210)
    $okButton.Add_Click({
        $controls["Start Date"].Text = $calendar.SelectionStart.ToString("yyyy-MM-dd")
        $calendarForm.Close()
    })
    $calendarForm.Controls.Add($okButton)

    $calendarForm.Topmost = $true
    $calendarForm.ShowDialog()
})

$buttons["End Date"].Add_Click({
    $calendarForm = New-Object System.Windows.Forms.Form
    $calendarForm.Text = "Select End Date"
    $calendarForm.Size = New-Object System.Drawing.Size(250, 280)
    $calendarForm.StartPosition = "CenterScreen"

    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.MaxSelectionCount = 1
    $calendar.Dock = "Top"
    $calendarForm.Controls.Add($calendar)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Location = New-Object System.Drawing.Point(80, 210)
    $okButton.Add_Click({
        $controls["End Date"].Text = $calendar.SelectionStart.ToString("yyyy-MM-dd")
        $calendarForm.Close()
    })
    $calendarForm.Controls.Add($okButton)

    $calendarForm.Topmost = $true
    $calendarForm.ShowDialog()
})

# Start Search button
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Start Search"
$btnSearch.Size = New-Object System.Drawing.Size(200, 30)
$btnSearch.Location = New-Object System.Drawing.Point(225, $y)

$btnSearch.Add_Click({
    # Collect field values
    $values = @{}
    foreach ($key in $controls.Keys) {
        if ($controls[$key] -is [System.Windows.Forms.TextBox] -or $controls[$key] -is [System.Windows.Forms.ComboBox]) {
            $values[$key] = $controls[$key].Text
        }
        elseif ($controls[$key] -is [System.Windows.Forms.CheckBox]) {
            $values[$key] = $controls[$key].Checked
        }
    }

    # Field Check
    $searchName = $values['Search Name']
    $senderEmail = $values['Sender Email']
    $scope = $values['Scope (subject/body)']
    $searchTerm = $values['Search Term']

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($searchName))   { $missing += "Search Name" }
    if ([string]::IsNullOrWhiteSpace($senderEmail) -or -not ($senderEmail -eq '*' -or $senderEmail -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or $senderEmail -match '^[^@\s]+\*$')) {
    $missing += "Sender Email (blank or invalid format)"}
    if ([string]::IsNullOrWhiteSpace($scope))        { $missing += "Scope" }
    if ([string]::IsNullOrWhiteSpace($searchTerm))   { $missing += "Search Term" }

    if ($missing.Count -gt 0) {
        $message = "The following mandatory fields are missing:`n`n" + ($missing -join "`n") + "`n`nPlease review and submit again."
        [System.Windows.Forms.MessageBox]::Show($message, "Missing Fields", 'OK', 'Error') | Out-Null
        return
    }

    # Build PowerShell Script
    Get-ChildItem -Path $env:TEMP -Filter "ComplianceSearch_*.ps1" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-24) } | Remove-Item -Force -ErrorAction SilentlyContinue
    $timestamp = Get-Date -Format "yyyy-MM-dd_hhmmtt"
    $tempPath = Join-Path $env:TEMP "ComplianceSearch_$timestamp.ps1"

    $script = @"
    #################################################################
    # Modules/Connection
    if (!(Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    Connect-ExchangeOnline
    if (!(Get-Command -Name Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Install-Module -Name ExchangeOnlineComplianceManagement -Scope CurrentUser -Force
    }
    Connect-IPPSSession -EnableSearchOnlySession

    #################################################################
    # Functions
        function Start-SearchSleepProgress {
    param([int]`$Num)
    1..`$Num | ForEach-Object {
        Write-Progress -Activity "Waiting..." -Status "Next status check in `$(`$Num - `$_)s" -PercentComplete (`$_ / `$Num * 100)
        Start-Sleep 1
    }
    Write-Progress -Activity "" -Completed
    }

    function Start-SleepProgress {
    param([int]`$Num)
    1..`$Num | ForEach-Object {
        Write-Progress -Activity "Sleeping for `$Num seconds" -Status "Remaining:`$(`$Num-`$_)" -PercentComplete (`$_/`$Num*100)
        Start-Sleep 1
    }
    Write-Progress -Activity "Sleeping for `$Num seconds" -Completed
    }

    #################################################################
    # Variables
    `$name = "$($values['Search Name'])"
    `$name = `$name.Substring(0, [Math]::Min(50, `$name.Length))
    `$fromemail = "$($values['Sender Email'])"
    `$searchScope = "$($values['Scope (subject/body)'])"
    `$searchTerm = "$($values['Search Term'])"
    `$startDate = "$($values['Start Date'])"
    `$endDate = "$($values['End Date'])"
    `$purgeSoft = "$($values['SoftDelete'])"
    `$purgeHard = "$($values['HardDelete'])"
    `$deleteSearch = "$($values['Delete Search'])"
    `$query = "Logic needed"

    
    #################################################################
    # Search - Setup
    Write-Host ("Compliance search started at " + (Get-Date -Format "MM/dd/yyyy hh:mm tt")) -ForegroundColor Green
    Write-Host "`n`nSearch Name: `$name"
    Write-Host "Sender Email: `$fromemail"
    Write-Host "Scope: `$searchScope"
    Write-Host "Search Term: `$(
        if (`$searchScope -eq 'subject' -and `$searchTerm -eq '*') {
            '* (Wildcard - All Messages)'
        } else {
            `$searchTerm
        })"
    
    if ("`$startDate" -ne "") { Write-Host "Start Date: `$startDate" } else { Write-Host "Start Date: Not Set" }
    if ("`$endDate" -ne "") { Write-Host "End Date: `$endDate" } else { Write-Host "End Date: Not Set" }
    
    if ("`$purgeSoft" -eq "True" -or "`$purgeHard" -eq "True") {
        `$purgeType = if ("`$purgeHard" -eq "True") { "HardDelete" } else { "SoftDelete" }
        Write-Host "Purge Type: `$purgeType"
    } else {
        Write-Host "Purge Type: None"
    }
    
    if ("`$deleteSearch" -eq "True") { Write-Host "Delete Search: Yes" } else { Write-Host "Delete Search: No" }
    
    if ("`$searchTerm" -eq "*") { `$searchTerm = `$null }
    if ("`$searchTerm" -match '^\*') { `$searchTerm = `$searchTerm.TrimStart('*') }
    if (`$searchTerm) { `$searchTerm = `$searchTerm.Trim() }
    
    if ("`$searchScope" -eq "subject") {
        if ("`$fromemail" -eq "*") {
            `$query = if ("`$searchTerm") {
                if ("`$startDate" -or "`$endDate") { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_')) (date=`$startDate..`$endDate)" } else { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_'))" }
            } else {
                if ("`$startDate" -or "`$endDate") { "(date=`$startDate..`$endDate)" } else { "*" }
            }
        } elseif ("`$fromemail" -match '\*') {
            `$query = if ("`$searchTerm") {
                if ("`$startDate" -or "`$endDate") { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_')) (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_')) (From:`$fromemail OR Participants:`$fromemail)" }
            } else {
                if ("`$startDate" -or "`$endDate") { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)" }
            }
        } else {
            `$query = if ("`$searchTerm") {
                if ("`$startDate" -or "`$endDate") { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_')) (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(Subject:`$searchTerm OR Subject:`$(`$searchTerm -replace ' ', '_')) (From:`$fromemail OR Participants:`$fromemail)" }
            } else {
                if ("`$startDate" -or "`$endDate") { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)" }
            }
        }
    } elseif ("`$searchScope" -eq "body") {
        if ("`$fromemail" -eq "*") {
            `$query = if (`$searchTerm) { "`$searchTerm (date=`$startDate..`$endDate)" } else { "(date=`$startDate..`$endDate)" }
        } elseif ("`$fromemail" -match '\*') {
            `$query = if (`$searchTerm) { "`$searchTerm (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" }
        } else {
            `$query = if (`$searchTerm) { "`$searchTerm (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" }
        }
    }
    
    #################################################################
    # Search - Start
    Write-Host "`nQuery: `$query" -ForegroundColor DarkYellow
    Write-Host "`nSearch Starting: `$name"
    
    # Timer - Start
    `$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    New-ComplianceSearch -Name `$name -ExchangeLocation All -ContentMatchQuery `$query | Out-Null
    Start-ComplianceSearch -Identity `$name
    
    do {
        Start-SearchSleepProgress -Num 60
        `$status = (Get-ComplianceSearch -Identity `$name).Status
        Write-Host "." -NoNewline
    } while (`$status -ne "Completed")
    
    #################################################################
    # Search - Results (Complete)
    # Timer - Stop
    `$stopwatch.Stop()
    `$totalTime = "{0:00}:{1:00}" -f `$stopwatch.Elapsed.Hours, `$stopwatch.Elapsed.Minutes
    Write-Host "`n - Search completed.`n - Time:`$totalTime"

    #################################################################
    # Search - Results (Display)
    `$search = Get-ComplianceSearch -Identity `$name -ErrorAction SilentlyContinue
    if (`$null -eq `$search) {
        Write-Host "Error: Unable to retrieve compliance search details. Please verify the search name." -ForegroundColor Red
        exit
    }

    `$items = `$search.Items
    Write-Host " - Items: '`$items'"
    
    #################################################################
    # Search - Results (Purge)
    if ("`$purgeHard" -eq "True") {
        `$type = "HardDelete"
    } elseif ("`$purgeSoft" -eq "True") {
        `$type = "SoftDelete"
    } else {
        `$type = `$null
    }

    if ("`$type" -eq "HardDelete" -or "`$type" -eq "SoftDelete") {
        Write-Host " - Purging: `$type"
        New-ComplianceSearchAction -SearchName `$name -Purge -PurgeType `$type -Confirm:`$false
        Start-SleepProgress -Num 300
    }

    if ("`$deleteSearch" -eq "True") {
        Remove-ComplianceSearch -Identity `$name -Confirm:`$false
        Write-Host "Search `$name was deleted.`n"
    }

    Write-Host ("`nCompliance search completed at " + (Get-Date -Format "MM/dd/yyyy hh:mm tt") + "`n") -ForegroundColor Green

    # Cleanup
    # Exchange Online
    `$null = Disconnect-ExchangeOnline -Confirm:`$false
"@
    
    $script | Set-Content -Path $tempPath -Encoding UTF8

    try {
        Start-Process -FilePath "pwsh.exe" -ArgumentList "-NoExit", "-File", "`"$tempPath`""
    } catch {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$tempPath`""
    }

    $form.Close()
})

$form.Controls.Add($btnSearch)
$form.ShowDialog()