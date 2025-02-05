function main() {
    # 起動済みのOutlookがあるか確認
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    $needQuit = $false
    if ($outlookProcess -eq $null) {
        $needQuit = $true
    }

    $outlook = New-Object -ComObject Outlook.Application
    try {
        $Color = [Microsoft.Office.Interop.Outlook.OlCategoryColor]
        $ShortCutKey = [Microsoft.Office.Interop.Outlook.OlCategoryShortCutKey]
        $namespace = $outlook.GetNamespace("MAPI")
        
        # 既存の全てのカテゴリを削除
        while ($namespace.Categories.Count -gt 0) {
            $namespace.Categories.Remove($namespace.Categories.Item(1).Name)
        }

        # カテゴリーの追加
        Start-Sleep -Seconds 1
        $namespace.Categories.Add("Client", $Color::olCategoryColorRed, $ShortCutKey::olCategoryShortcutKeyCtrlF2)
        Start-Sleep -Seconds 1
        $namespace.Categories.Add("Personal", $Color::olCategoryColorGray, $ShortCutKey::olCategoryShortcutKeyCtrlF3)
        Start-Sleep -Seconds 1
        $namespace.Categories.Add("Block", $Color::olCategoryColorBlack, $ShortCutKey::olCategoryShortcutKeyCtrlF4)
        
        # 土日にOOOを設定
        $event_item = $outlook.CreateItem(1) # 1 = olAppointmentItem
        $event_item.Subject = "Holiday"
        $event_item.Location = "Japan"
        $event_item.AllDayEvent = $true
        $event_item.BusyStatus = 3
        $event_item.ReminderSet = $false
        
        # Set recurrence pattern
        $pattern = $event_item.GetRecurrencePattern()
        $pattern.RecurrenceType = 1
        $pattern.DayOfWeekMask = [Microsoft.Office.Interop.Outlook.OlDaysOfWeek]::olSunday + [Microsoft.Office.Interop.Outlook.OlDaysOfWeek]::olSaturday
        $pattern.NoEndDate = $true
        
        $event_item.Save()
    }
    finally {
        if ($needQuit) {
            [void]$outlook.Quit()
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
        }
    }
}

main