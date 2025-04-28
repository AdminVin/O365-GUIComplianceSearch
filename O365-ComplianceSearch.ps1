# GUI Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Compliance Search"
$form.Size = New-Object System.Drawing.Size(400, 450)
$form.StartPosition = "CenterScreen"

$labels = @("Search Name", "Sender Email (* = any)", "Scope (subject/body)", "Search Term", "Start Date", "End Date")
$textboxes = @{}
$y = 10

foreach ($label in $labels) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $label
    $lbl.Location = New-Object System.Drawing.Point(10, $y)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Size = New-Object System.Drawing.Size(250, 20)
    $txt.Location = New-Object System.Drawing.Point(120, $y)
    $form.Controls.Add($txt)
    $textboxes[$label] = $txt

    $y += 40
}

$btnStartDate = New-Object System.Windows.Forms.Button
$btnStartDate.Text = "Pick Start"
$btnStartDate.Size = New-Object System.Drawing.Size(80, 20)
$btnStartDate.Location = New-Object System.Drawing.Point(300, 170)
$btnStartDate.Add_Click({
    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.MaxSelectionCount = 1
    $calendar.ShowTodayCircle = $false
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Select Start Date"
    $popup.Size = New-Object System.Drawing.Size(250, 250)
    $calendar.Dock = "Fill"
    $popup.Controls.Add($calendar)
    $popup.Topmost = $true
    $popup.ShowDialog()
    $textboxes["Start Date"].Text = $calendar.SelectionStart.ToString("yyyy-MM-dd")
})
$form.Controls.Add($btnStartDate)

$btnEndDate = New-Object System.Windows.Forms.Button
$btnEndDate.Text = "Pick End"
$btnEndDate.Size = New-Object System.Drawing.Size(80, 20)
$btnEndDate.Location = New-Object System.Drawing.Point(300, 210)
$btnEndDate.Add_Click({
    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.MaxSelectionCount = 1
    $calendar.ShowTodayCircle = $false
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Select End Date"
    $popup.Size = New-Object System.Drawing.Size(250, 250)
    $calendar.Dock = "Fill"
    $popup.Controls.Add($calendar)
    $popup.Topmost = $true
    $popup.ShowDialog()
    $textboxes["End Date"].Text = $calendar.SelectionStart.ToString("yyyy-MM-dd")
})
$form.Controls.Add($btnEndDate)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Start Search"
$btnSearch.Size = New-Object System.Drawing.Size(360, 30)
$btnSearch.Location = New-Object System.Drawing.Point(10, 300)
$btnSearch.Add_Click({
    $name = $textboxes["Search Name"].Text.Trim()
    $fromemail = $textboxes["Sender Email (* = any)"].Text.Trim()
    $scope = $textboxes["Scope (subject/body)"].Text.Trim().ToLower()
    $term = $textboxes["Search Term"].Text.Trim()
    $startDate = $textboxes["Start Date"].Text.Trim()
    $endDate = $textboxes["End Date"].Text.Trim()

    if (!$name -or !$fromemail -or !$scope -or !$startDate -or !$endDate) {
        [System.Windows.Forms.MessageBox]::Show("Please fill all required fields.","Error",'OK','Error')
        return
    }

    if (!(Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        Connect-ExchangeOnline
    } else {
        Connect-ExchangeOnline
    }

    if (!(Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Install-Module ExchangeOnlineComplianceManagement -Scope CurrentUser -Force
        Connect-IPPSSession
    } else {
        Connect-IPPSSession
    }

    if ($term -eq "*") { $term = $null }
    if ($term -match '^\*') { $term = $term.TrimStart('*') }

    if ($scope -eq "subject") {
        if ($fromemail -eq "*") {
            $query = $term ? "(Subject:$term) (date=$startDate..$endDate)" : "(date=$startDate..$endDate)"
        } elseif ($fromemail -match '\*') {
            $query = $term ? "(Subject:$term) (From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)" : "(From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)"
        } else {
            $query = $term ? "(Subject:$term) (From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)" : "(From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)"
        }
    } elseif ($scope -eq "body") {
        if ($fromemail -eq "*") {
            $query = $term ? "$term (date=$startDate..$endDate)" : "(date=$startDate..$endDate)"
        } elseif ($fromemail -match '\*') {
            $query = $term ? "$term (From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)" : "(From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)"
        } else {
            $query = $term ? "$term (From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)" : "(From:$fromemail OR Participants:$fromemail)(date=$startDate..$endDate)"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Scope must be 'subject' or 'body'.","Error",'OK','Error')
        return
    }

    Write-Host "Search Query: $query" -ForegroundColor Green

    New-ComplianceSearch -Name $name -ExchangeLocation "All" -ContentMatchQuery $query | Out-Null
    Start-ComplianceSearch -Identity $name

    Write-Host "Searching..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ((Get-ComplianceSearch $name -ErrorAction SilentlyContinue).Status -ne "Completed") {
        Start-Sleep -Seconds 1
    }
    $stopwatch.Stop()

    Write-Host "Search Completed in $($stopwatch.Elapsed.Minutes) minutes."
    
    $search = Get-ComplianceSearch -Identity $name -ErrorAction SilentlyContinue
    $items = $search.Items
    $results = $search.SuccessResults
    $mailboxes = @()
    if ($results -is [string] -and $results -ne "") {
        foreach ($line in $results -split '[\r\n]+') {
            if ($line -match 'Location: (\S+),.+Item count: (\d+)' -and $matches[2] -gt 0) {
                $mailboxes += $matches[1]
            }
        }
    }
    Write-Host "Mailboxes:"
    $mailboxes
    Write-Host "Total items found: '$items'"

    $purgePrompt = Read-Host "Type 'purge' to delete items, or press Enter to exit"
    if ($purgePrompt -eq "purge") {
        New-ComplianceSearchAction -SearchName $name -Purge -PurgeType SoftDelete -Confirm:$false
        Write-Host "Purging..."
        Start-Sleep -Seconds 300
        $deletePrompt = Read-Host "Delete compliance search after purge? (Y/N)"
        if ($deletePrompt -eq "Y") {
            Remove-ComplianceSearch -Identity $name -Confirm:$false
            Write-Host "ComplianceSearch deleted."
        }
    } else {
        $deletePrompt = Read-Host "Delete compliance search? (Y/N)"
        if ($deletePrompt -eq "Y") {
            Remove-ComplianceSearch -Identity $name -Confirm:$false
            Write-Host "ComplianceSearch deleted."
        }
    }

    Get-PSSession | Remove-PSSession | Out-Null
})
$form.Controls.Add($btnSearch)

$form.ShowDialog()
