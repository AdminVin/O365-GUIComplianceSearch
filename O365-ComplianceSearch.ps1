# GUI Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Compliance Search"
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
        $note.Text = "Use * for searching all messages."
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

    # Build PowerShell script content
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempPath = Join-Path $env:TEMP "ComplianceSearch_$timestamp.ps1"

    $script = @"
    if (!(Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    Connect-ExchangeOnline
    if (!(Get-Command -Name Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Install-Module -Name ExchangeOnlineComplianceManagement -Scope CurrentUser -Force
    }
    Connect-IPPSSession
    
    `$name = "$($values['Search Name'])"
    `$fromemail = "$($values['Sender Email'])"
    `$searchScope = "$($values['Scope (subject/body)'])"
    `$searchTerm = "$($values['Search Term'])"
    `$startDate = "$($values['Start Date'])"
    `$endDate = "$($values['End Date'])"
    `$purgeSoft = "$($values['SoftDelete'])"
    `$purgeHard = "$($values['HardDelete'])"
    `$deleteSearch = "$($values['Delete Search'])"
    
    Write-Host ("Compliance search started at " + (Get-Date -Format "MM/dd/yyyy hh:mm tt")) -ForegroundColor Green
    Write-Host "`n`nSearch Name: `$name" -ForegroundColor White
    Write-Host "Sender Email: `$fromemail" -ForegroundColor White
    Write-Host "Scope: `$searchScope" -ForegroundColor White
    Write-Host "Search Term: `$(
        if (`$searchScope -eq 'subject' -and `$searchTerm -eq '*') {
            '* (Wildcard - All Messages)'
        } else {
            `$searchTerm
        })" -ForegroundColor White
    
    if ("`$startDate" -ne "") { Write-Host "Start Date: `$startDate" -ForegroundColor White } else { Write-Host "Start Date: Not Set" -ForegroundColor White }
    if ("`$endDate" -ne "") { Write-Host "End Date: `$endDate" -ForegroundColor White } else { Write-Host "End Date: Not Set" -ForegroundColor White }
    
    if ("`$purgeSoft" -eq "True" -or "`$purgeHard" -eq "True") {
        `$purgeType = if ("`$purgeHard" -eq "True") { "HardDelete" } else { "SoftDelete" }
        Write-Host "Purge Type: `$purgeType" -ForegroundColor White
    } else {
        Write-Host "Purge Type: None" -ForegroundColor White
    }
    
    if ("`$deleteSearch" -eq "True") { Write-Host "Delete Search: Yes" -ForegroundColor White } else { Write-Host "Delete Search: No" -ForegroundColor White }
    
    if ("`$searchTerm" -eq "*") { `$searchTerm = `$null }
    if ("`$searchTerm" -match '^\*') { `$searchTerm = `$searchTerm.TrimStart('*') }
    
    if ("`$searchScope" -eq "subject") {
        if ("`$fromemail" -eq "*") {
            `$query = if (`$searchTerm) { "(Subject:`$searchTerm) (date=`$startDate..`$endDate)" } else { "(date=`$startDate..`$endDate)" }
        } elseif ("`$fromemail" -match '\*') {
            `$query = if (`$searchTerm) { "(Subject:`$searchTerm) (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" }
        } else {
            `$query = if (`$searchTerm) { "(Subject:`$searchTerm) (From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" } else { "(From:`$fromemail OR Participants:`$fromemail)(date=`$startDate..`$endDate)" }
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
    
    Write-Host "`nQuery: `$query`n" -ForegroundColor DarkYellow
    
    New-ComplianceSearch -Name `$name -ExchangeLocation All -ContentMatchQuery `$query | Out-Null
    Start-ComplianceSearch -Identity `$name
    
    do {
        Start-Sleep 1
        `$status = (Get-ComplianceSearch -Identity `$name).Status
        Write-Host "." -NoNewline
    } while (`$status -ne "Completed")
    
    Write-Host "`nSearch complete.`n"
    
    if ("`$purgeSoft" -eq "True" -or "`$purgeHard" -eq "True") {
        `$type = if ("`$purgeHard" -eq "True") { "HardDelete" } else { "SoftDelete" }
        Write-Host "`nPurging via `$type..."
        New-ComplianceSearchAction -SearchName `$name -Purge -PurgeType `$type -Confirm:`$false
        Start-Sleep -Seconds 5
        if ("`$deleteSearch" -eq "True") {
            Remove-ComplianceSearch -Identity `$name -Confirm:`$false
            Write-Host "`nSearch deleted."
        }
    }
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
