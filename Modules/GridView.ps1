add-type -AssemblyName 'System.Windows.Forms'
Add-Type -AssemblyName 'System.Drawing'
Add-Type -AssemblyName 'PresentationFramework'
$sys = [System.Environment]::SystemDirectory

$selection = [System.Collections.ArrayList]::new()
$array = [System.Collections.ArrayList]::new()
$day_start = [datetime]::Today
$day_end = [datetime]::Today.AddDays(1).AddSeconds(-1)

add-type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlookFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$outlook = New-Object -ComObject outlook.application
$namespace = $outlook.GetNamespace("MAPI")
$folder = $namespace.GetDefaultFolder($outlookFolders::olFolderCalendar)
$items = $folder.Items
$meetings = $items|?{$_.Start -ge $day_start -and $_.End -le $day_end }|Select Subject, @{l='Meeting Start';e={$_.Start.toString("HH:mm tt")}}, @{l='Mask';e={$_.GetRecurrencePattern().DayOfWeekMask}}

$main2 = New-Object System.Windows.Forms.Form
$main2.Icon = $icon3
$main2.Size = '400,310'
$main2.MinimizeBox = $false
$main2.MaximizeBox = $false
$main2.Text = 'Meets'  
$main2.AutoSize = $true 






$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = '5,5'
$grid.Height = 300
$grid.Width  = 390
$grid.ColumnHeadersVisible = $true
$grid.AutoSizeColumnsMode = 'AllCells'
$grid.MultiSelect = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSize = $true
$grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$array.AddRange(@($meetings))
$grid.DataSource=($array)
$dataBindingsComplete = {
param([object]$sender, [System.EventArgs]$e)
$grid.Columns[2].Visible = $false
}


$button = [System.Windows.Forms.Button]::new()
$button.Location = [System.Drawing.Point]::new(5,310)
$button.Size = [System.Drawing.Size]::new(100,50)
$button.Text = 'OK'
$button.Add_Click({

$grid.SelectedRows|%{
$Script:selection.Add([pscustomobject]@{
                           Name      = $grid.Rows[$_.Index].Cells[0].Value
                           Time      = $grid.Rows[$_.Index].Cells[1].Value
                           DayofWeek = $grid.Rows[$_.Index].Cells[2].Value

})
}
$Script:selection|%{$task_name = $_.Name
#$task_name|out-host
#$taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $Script:task_name }
#$taskExists|Out-Host
#if(!$taskExists){
$new_hour = ([datetime]$_.Time).addminutes(-2)
$trigger = New-ScheduledTaskTrigger -At $new_hour -Once
$action = New-ScheduledTaskAction -Execute "$sys\wscript.exe" -Argument "//nologo $env:Userprofile\$task_name.vbs"
Register-ScheduledTask -TaskName $task_name -Trigger $trigger -Action $action 
$task = Get-ScheduledTask -TaskName $task_name
$task.Triggers[0].EndBoundary = $new_hour.AddMinutes(1).ToString('s')
$task.Settings.DeleteExpiredTaskAfter = 'PT0S'
$task|Set-ScheduledTask
 



$asd = "'" + $_.Name + "'"

$name = $ExecutionContext.InvokeCommand.ExpandString($asd)

$MyScript = @'
`$shell = New-Object -ComObject "Shell.Application"
`$shell.minimizeall()

add-type -AssemblyName 'System.Windows.Forms'
`$main2 = New-Object System.Windows.Forms.Form
`$main2.ClientSize = '400,300'
`$main2.MinimizeBox = `$false
`$main2.MaximizeBox = `$false
`$main2.Text = 'Przypominajka'
`$main2.AutoSize = `$true 
`$main2.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen 
`$label1 = [System.Windows.Forms.Label]::new()
`$label1.Dock = [System.Windows.Forms.DockStyle]::Fill
`$label1.Text = $name + '`r`n za 2 minuty!'
`$label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
`$label1.Font = [System.Drawing.Font]::new('Arial','25', [System.Drawing.FontStyle]::Underline)
`$label1.ForeColor = [System.Drawing.Color]::OrangeRed
`$timer = New-Object System.Windows.Forms.Timer
`$timer.Interval = 250
`$timer.Add_Tick({`$label1.Visible = -not(`$label1.Visible)})
`$main2.Add_Load({`$timer.Start()})
`$main2.Controls.Add(`$label1)
`$main2.ShowDialog()
`$timer.Dispose()
`$main2.Dispose()
'@

$2 = $ExecutionContext.InvokeCommand.ExpandString($MyScript)

$MyEncodedScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($2))

$a ='Set objFSO = CreateObject("Scripting.FileSystemObject")' + "`r`n" + 'strScript = Wscript.ScriptFullName' + "`r`n" + 'objFSO.DeleteFile(strScript)' +"`r`n" + 'dim EncodedCommand' + "`r`n" + 'EncodedCommand = ' + '"' + $MyEncodedScript + '"' + "`r`n" + 'pSCmd = "powershell.exe -noexit -windowstyle Hidden -executionpolicy bypass -encodedcommand " & EncodedCommand' + "`r`n" + 'CreateObject("WScript.Shell").Run pSCmd, 0, True'

new-item -Path $env:USERPROFILE -Name ($task_name + '.vbs') -Value $a

}
$Script:selection.Clear()

})

$cancel = [System.Windows.Forms.Button]::new()
$cancel.Location = [System.Drawing.Point]::new(295 ,310)
$cancel.Size = [System.Drawing.Size]::new(100,50)
$cancel.Text = 'CLOSE'

$cancel.Add_Click({
$main2.Close()
$main2.Dispose()
})

$main2.Add_Load($dataBindingsComplete)
$main2.Controls.AddRange(@($grid,$button, $cancel))
$main2.ShowDialog()
