# Workstation User Query
# Written by Joshua Woleben
# Written on 10/22/2019

$functions = @'
function write_log {
    Param([string]$log_entry,
            [string]$TranscriptFile)

            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                
            }
            finally {
                $mutex.ReleaseMutex() | out-null

            }
            $mutex.Dispose()
}
'@
function write_log {
    Param([string]$log_entry,
            [string]$TranscriptFile)

            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                
            }
            finally {
                $mutex.ReleaseMutex() | out-null

            }
            $mutex.Dispose()
}
$UserFunction = {
    Param([string]$functions,[string]$workstation,[string]$status_file)
    Invoke-Expression $functions
    if (Test-Connection -ComputerName $workstation -Count 1 -Quiet) {
        $session = New-PSSession -ComputerName $workstation
        $user_array =@()
        # $user = (Get-WmiObject -Class win32_loggedonuser -ComputerName $workstation | Where {$_.Path -notmatch "Domain.*$workstation" } | Select -ExpandProperty __RELPATH | ForEach-Object { (Select-String -InputObject $_.ToString() -Pattern "Name=(.*?),").Matches.Groups[1].Value  } | Where {$_ -notmatch "$env:USERNAME|is\.service"} | Select -Unique) -replace "\\`"",'' -replace "`"","" -join " "
        $user_list = $(Invoke-Command -Session $session -ScriptBlock {
            & query user
        })
        Remove-PSSession -Session $session
        for ($i=1;$i -lt $user_list.Count; $i++) {
           $user_array += ($user_list[$i] | Select-String -Pattern "\s(.+?)\s+").Matches.Groups[1].Value -replace "is\.service",""
        }
        $user = [string]::join(" ",$user_array)

    }
    else {
        $user = "Powered off"
    }
    if ([string]::IsNullOrEmpty($user)) {
        $user = "No user logged in"
    }
    $message = ($workstation + ", " + $user) 
    write_log $message $status_file

}
# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Workstation User Query" Height="1000" Width="800" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="FileLabel" Content="File to Search"/>
        <TextBox x:Name="FileTextBox" Height="20"/>
        <Button x:Name="FileSearchButton" Content="Query Workstations" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Button x:Name="ClearFormButton" Content="Clear Form" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Label x:Name="ResultsLabel" Content="Workstation User Results"/>
        <DataGrid x:Name="Results" AutoGenerateColumns="True" Height="400">
            <DataGrid.Columns>
                <DataGridTextColumn Header="WorkstationName" Binding="{Binding WorkstationName}" Width="200"/>
                <DataGridTextColumn Header="LoggedInUser" Binding="{Binding LoggedInUser}" Width="450"/>
            </DataGrid.Columns>
        </DataGrid>
        <Button x:Name="ExcelButton" Content="Export to Excel" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$FileTextBox = $global:Form.FindName('FileTextBox')
$FileSearchButton = $global:Form.FindName('FileSearchButton')
$Results = $global:Form.FindName('Results')
$ClearFormButton = $global:Form.FindName('ClearFormButton')
$ExcelButton = $global:Form.FindName('ExcelButton')

$FileSearchButton.Add_Click({

    $status_file = "C:\Temp\status_file.txt"
    if (Test-Path -Path $status_file) {
        Remove-Item -Path $status_file
    }
    $file_to_load = $FileTextBox.Text

    Write-Host "Loading file $file_to_load..."

    $workstation_list = Get-Content $file_to_load

    Write-Host "Starting workstation query..."
    foreach ($workstation in $workstation_list) {
        Write-Host "Querying workstation $workstation..."
        Start-Job -ArgumentList @($functions,$workstation,$status_file) -ScriptBlock $UserFunction
        while (@(Get-Job).Count -gt 100) {

            Write-Host "." -NoNewline
                        
            Remove-Job -State Completed
            sleep 5
        }
    }
    Write-Host "Waiting for jobs..."
    Get-Job | Wait-Job -Timeout 180
    $final_output = Get-Job | Receive-Job
    $final_output = $(($final_output | Out-String) -join "`n")
    Get-Job | Remove-Job -Force

    Write-Host "Adding results..."
    $final_status = Get-Content $status_file 
    foreach ($line in $final_status) {
        $workstation = (Select-String -Pattern "(.*),.*" -InputObject $line).Matches.Groups[1].Value
        $user = (Select-String -Pattern ".*,(.*)" -InputObject $line).Matches.Groups[1].Value
        $Results.AddChild([PSCustomObject]@{WorkstationName=$workstation; LoggedInUser=$user})
    }
    [System.Windows.MessageBox]::Show("Results loaded!")
})

$ClearFormButton.Add_Click({
    $FileSearchTextBox.Text = ""
    $Results.Items.Clear()
    $global:Form.invalidateVisual()
})
$ExcelButton.Add_Click({
    $excel_file = "$env:USERPROFILE\Documents\LoggedInUserQuery_$(get-date -f MMddyyyyHHmmss).xlsx"
    # Open Excel

    # Create new Excel object
    $excel_object = New-Object -comobject Excel.Application
    $excel_object.visible = $True 

    # Create new Excel workbook
    $excel_workbook = $excel_object.Workbooks.Add()

    # Select the first worksheet in the new workbook
    $excel_worksheet = $excel_workbook.Worksheets.Item(1)

    # Create headers
    $excel_worksheet.Cells.Item(1,1) = "Workstation Name"
    $excel_worksheet.Cells.Item(1,2) = "Logged in Users"

    # Format headers
    $d = $excel_worksheet.UsedRange

    # Set headers to backrgound pale yellow color, bold font, blue font color
    $d.Interior.ColorIndex = 19
    $d.Font.ColorIndex = 11
    $d.Font.Bold = $True

    # Set first data row
    $row_counter = 2
    Foreach ($item in $Results.Items) {
        $excel_worksheet.Cells.Item($row_counter,1) = $item.WorkstationName
        $excel_worksheet.Cells.Item($row_counter,2) = $item.LoggedInUser
        $row_counter++
    }
    # Create borders around the cell in the spreadsheet. The below code creates all borders
    $e = $excel_worksheet.Range("A1:B$row_counter")
    $e.Borders.Item(12).Weight = 2
    $e.Borders.Item(12).LineStyle = 1
    $e.Borders.Item(12).ColorIndex = 1

    $e.Borders.Item(11).Weight = 2
    $e.Borders.Item(11).LineStyle = 1
    $e.Borders.Item(11).ColorIndex = 1

    # Set thicker border around outside
    $e.BorderAround(1,4,1)

    # Fit all columns
    $e.Columns("A:B").AutoFit()

    # Save Excel
    $excel_workbook.SaveAs($excel_file) | out-null

    # Quit Excel
   # $excel_workbook.Close | out-null
   # $excel_object.Quit() | out-null

    [System.Windows.MessageBox]::Show("File written to $excel_file")
})

# Show GUI
$global:Form.ShowDialog() | out-null