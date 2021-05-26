[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[reflection.assembly]::LoadWithPartialName("System.Drawing") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

#Create a WinForm
$folderform = New-Object System.Windows.Forms.Form
$folderform.Size = New-Object System.Drawing.Size(600, 700)
$folderform.BackColor = 'Gray'
$folderform.Text = 'File Cap'
$folderform.StartPosition = 'CenterScreen'
$folderform.FormBorderStyle = 'FixedDialog'
$text_font = 'Time New Roman, 16.0'

#Database login credentials       
$sqlserver = "yourserver"
$db = "yourdb"
$usrname = "yourusername"
$pwd = "yourpwd"

#First segment: select folder or type one
$folder_label = New-Object System.Windows.Forms.Label
$folder_label.Location = New-Object System.Drawing.Size(20, 40)
$folder_label.Size = New-Object System.Drawing.Size(130, 30)
$folder_label.Text = 'Folder Path'
<# Cmd for auto-resize
$folder_label.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
-bor [System.Windows.Forms.AnchorStyles]::Bottom `
-bor [System.Windows.Forms.AnchorStyles]::Left `
-bor [System.Windows.Forms.AnchorStyles]::Right#>

$filepath = New-Object System.Windows.Forms.TextBox
$filepath.Location = '20, 80'
$filepath.Size = '400, 30'
$filepath.Multiline = $true


$select_button = New-Object System.Windows.Forms.Button
$select_button.Location = '440, 80'
$select_button.Size = '90, 30'
$select_button.Text = 'Select'
$select_button.BackColor = 'Lightblue'


$folder_brow = New-Object System.Windows.Forms.FolderBrowserDialog
$select_button.Add_Click({
    $folder_brow.ShowDialog()
    $filepath.Text = $folder_brow.SelectedPath
    $filepath.Font = $text_font
})

#Second segment: Get the number of files contained
$File_Count_label = New-Object System.Windows.Forms.Label
$File_Count_label.Location = '20, 160'
$File_Count_label.Size = '380, 30'
$File_Count_label.Text = 'Number of Files in Selected Folder'

$File_Count = New-Object System.Windows.Forms.TextBox
$File_Count.Location = '20, 200'
$File_Count.Size = '150, 30'
$File_Count.Multiline = $true

$File_Count_button = New-Object System.Windows.Forms.Button
$File_Count_button.Location = '200, 200'
$File_Count_button.Size = '170, 30'
$File_Count_button.BackColor = 'Lightblue'
$File_Count_button.Text = 'Get File Count'
$File_Count_button.Add_Click({
    $File_Count.Font = $text_font
    $File_Count.Text = (Get-ChildItem -Path $filepath.Text -File | Measure-Object).Count
})

#Third segment: Fetch the current file cap
$Current_Cap_label = New-Object System.Windows.Forms.Label
$Current_Cap_label.Location = '20, 280'
$Current_Cap_label.Size = '150, 30'
$Current_Cap_label.Text = 'Current Cap'

$Current_Cap = New-Object System.Windows.Forms.TextBox
$Current_Cap.Location = '20, 320'
$Current_Cap.Size = '150, 30'
$Current_Cap.Multiline = $true

#warning msg that only displays when no cap is found
$No_Cap_label = New-Object System.Windows.Forms.Label
$No_Cap_label.Location = '20, 350'
$No_Cap_label.Size = '300, 20'

$Current_Cap_button = New-Object System.Windows.Forms.Button
$Current_Cap_button.Location = '200, 320'
$Current_Cap_button.Size = '150, 30'
$Current_Cap_button.Text = 'Get File Cap'
$Current_Cap_button.BackColor = 'Lightblue'
$Current_Cap_button.Add_Click({
    if ($filepath.Text) {
        $filepath_text = $filepath.Text
        try {
            $connString = "Data Source=$sqlserver; Database=$db; User ID=$usrname; Password=$pwd"
            $conn = New-Object System.Data.SqlClient.SqlConnection $connString
            $conn.Open()
        
            if ($conn.State -eq 'Open') {
                Write-Host "Connection succeeded."
                $sqlcmd = $conn.CreateCommand()
                $sqlcmd.CommandText = "SELECT CAP FROM File_Cap
                                       WHERE FILEPATH= '$filepath_text' "
                
                $scalar = $sqlcmd.ExecuteScalar()
                
                if ($scalar -eq $null ) {
                    $No_Cap_label.Text = 'No cap found, set a new one!'
                    $No_Cap_label.ForeColor = 'Red'
                    }
                 else {
                    $Current_Cap.Font = $text_font
                    $Current_Cap.Text = $scalar
                    $File_Count_text = $File_count.Text 
                     
                    if ($scalar -lt [int]$File_Count_text) {  
                        $msg = "Current cap $scalar exceeded, Wanna upload extra files to cloud?"
                        $title = "File Excess"
                        $Button = "YesNo"
                        $Icon = "Warning"
                        $msgBoxInput = [System.Windows.MessageBox]::Show($msg, $title, $Button, $Icon)

                        #uploads extra files to the cloud if file cap is exceeded
                        switch ($msgBoxInput) {
                            #Launch IE if user agrees, could also use Edge with a webdriver
                            "Yes" {
                                $IE = New-Object -ComObject InternetExplorer.Application
                                $IE.Navigate("https://pan.baidu.com")
                                $IE.Visible=$true
                                [System.Threading.Thread]::Sleep(4000)   #wait until the page is loaded
                                $doc = $IE.Document
                                $userid = $doc.GetElementById("TANGRAM__PSP_4__userName")
                                $pwd = $doc.GetElementById("TANGRAM__PSP_4__password")
                                $userid.value = "yourusername"
                                $pwd.value = "yourpwd"
                                $btSubmit = $doc.getElementById("TANGRAM__PSP_4__submit")
                                $btSubmit.click()
                            }
                                "No" {"Action aborted"; exit}
                                                }
                      }
                    }
                }
            }
         catch {
            Write-Host "Connection failed."
        }
}
})

#Fourth segment; Set a new cap
$New_Cap_label = New-Object System.Windows.Forms.Label
$New_Cap_label.Location = '20, 400'
$New_Cap_label.Size = '200, 30'
$New_Cap_label.Text = "Set A New Cap"

$New_Cap = New-Object System.Windows.Forms.TextBox
$New_Cap.Location = '20, 440'
$New_Cap.Size = '150, 30'
$New_Cap.Multiline = $true
$New_Cap.Font = $text_font

#Feedback msg after a new cap is successfully set
$New_Cap_Set_label = New-Object System.Windows.Forms.Label
$New_Cap_Set_label.Location = '20, 480'
$New_Cap_Set_label.Size = '380, 30'


$New_Cap_button = New-Object System.Windows.Forms.Button
$New_Cap_button.Location = '200, 440'
$New_Cap_button.Size = '100, 30'
$New_Cap_button.Text = 'Update'
$New_Cap_button.BackColor = 'Lightblue'
$New_Cap_button.Add_Click({
    if ( ([int]$New_Cap.Text) -gt 0) {
        $New_Cap_text = $New_Cap.Text
        $filepath_text = $filepath.Text
        try { 
            $connString = "Data Source=$sqlserver; Database=$db; User ID=$usrname; Password=$pwd"
            $conn = New-Object System.Data.SqlClient.SqlConnection $connString
            $conn.Open()

            if ($conn.State -eq 'Open') {
                Write-Host "Database ready."
                if ($Current_Cap.Text) {
                    $sqlcmd_update = $conn.CreateCommand()
                    $sqlcmd_update.CommandText = "UPDATE File_Cap
                                                  SET CAP=$New_Cap_text
                                                  WHERE FILEPATH= '$filepath_text' "
                    $sqlcmd_update.ExecuteNonQuery()
                    $New_Cap_Set_label.Text = 'Cool! The file cap has been updated.'
                    $New_Cap_Set_label.ForeColor = 'Lightgreen'
                } else {
                    $sqlcmd_insert = $conn.CreateCommand()
                    $sqlcmd_insert.CommandText = "INSERT INTO File_Cap (FILEPATH, CAP)
                                                  VALUES ('$filepath_text', $New_Cap_text)"
                    $sqlcmd_insert.ExecuteNonQuery()

                    $New_Cap_Set_label.Text = 'Great! A new cap is set.'
                    $New_Cap_Set_label.ForeColor = 'Green'
                }
            }
        } catch {Write-Host "Failed to open a connection."}
    } else {
        $New_Cap_Set_label.Text = 'Please enter a valid file cap(1-100).'
        $New_Cap_Set_label.ForeColor = 'Red'
    }
}) 


#Create concluding okay and cancel button
$ok_button = New-Object System.Windows.Forms.Button
$ok_button.Location = '100, 550'
$ok_button.Size = '90, 30'
$ok_button.Text = 'OKAY'
$ok_button.Add_Click({
    $folderform.Controls | Where-Object{$_ -is [System.Windows.Forms.TextBox]} | ForEach-Object{$_.Clear()}
})
$ok_button.ForeColor = 'Black'



$cancel_button = New-Object System.Windows.Forms.Button
$cancel_button.Location = '300, 550'
$cancel_button.Size = '130, 30'
$cancel_button.Text = 'CANCEL'
$cancel_button.ForeColor = 'Red'

$folderform.AcceptButton = $ok_button
$folderform.CancelButton = $cancel_button

#Add all controls
$folderform.Controls.Add($folder_label)
$folderform.Controls.Add($filepath)
$folderform.Controls.Add($select_button)
$folderform.Controls.Add($Current_Cap_label)
$folderform.Controls.Add($Current_Cap)
$folderform.Controls.Add($Current_Cap_button)
$folderform.Controls.Add($New_Cap_label)
$folderform.Controls.Add($New_Cap)
$folderform.Controls.Add($New_Cap_button)
$folderform.Controls.Add($ok_button)
$folderform.Controls.Add($cancel_button)
$folderform.Controls.Add($File_Count_label)
$folderform.Controls.Add($File_Count)
$folderform.Controls.Add($File_Count_button)
$folderform.Controls.Add($No_Cap_label)
$folderform.Controls.Add($New_Cap_Set_label)
$folderform.ShowDialog()



