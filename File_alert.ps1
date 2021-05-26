$filepath = Read-Host -Prompt "Full path to the folder you want to probe"
$filecount = (Get-ChildItem -Path $filepath -File | Measure-Object).Count

try {
    $connString = "Data Source=$sqlserver; Database=$db; User ID=$usrname; Password=$pwd"
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    $conn.Open()

    if ($conn.State -eq 'Open') {
        Write-Host "Connection succeeded."
        $sqlcmd = $conn.CreateCommand()
        $sqlcmd.CommandText = "SELECT CAP FROM File_Cap
                               WHERE FILEPATH='$filepath' "

        $scalar = $sqlcmd.ExecuteScalar()

        if ($scalar -eq $null ) {
            $set_cap = Read-Host -Prompt "Current cap is null, would you like to set one(Yes/No)"
            switch ($set_cap) {
                "Yes" {
                    $newcap = Read-Host -Prompt "New cap(1-100)"
                    $sql_cmd_create = $conn.CreateCommand()
                    $sql_cmd_create.CommandText = "INSERT INTO File_Cap (FILEPATH, CAP)
                                                   VALUES ('$filepath', $newcap)"
                    $sql_cmd_create.ExecuteNonQuery()
                    $global:filelimit = $newcap
                    break
                 }
                 "No" {Write-Host "Exiting..."; exit}
            }
        } else {
            $choice = Read-Host -Prompt "Current cap is $scalar, want to modify it?(Yes/No)"
            switch ($choice) {
                "Yes" { 
                       $updatecap = Read-Host -Prompt "Set a new limit(1-100)"
                       if ($updatecap -le 0) {
                                    Write-Host "Invalid input, exiting without setting a new cap."
                                    exit
                                            }
                       else {
                                    $updatesql = "UPDATE File_Cap
                                                  SET CAP=$updatecap
                                                  WHERE FILEPATH='$filepath'"
                                    $sql_cmd_update = $conn.CreateCommand()
                                    $sql_cmd_update.CommandText = $updatesql
                                    $sql_cmd_update.ExecuteNonQuery()
                                    $msg2 = "File cap is now {0}." -f $updatecap
                                    Write-Host $msg2
                                    break
                             }      #update the related record in File_Cap
                       }
                 "No" {$global:filelimit = $scalar; break}
             }
         }
    }
}
catch {
    Write-Host "Connection failed!"
}

if ( $filecount -gt $filelimit)
{   
    #Create a interactive message box asking for user confirmation
    $msg = "Current cap $filelimit exceeded, Wanna upload extra files to cloud?"
    $title = "File Excess"
    $Button = "YesNo"
    $Icon = "Warning"
    $msgBoxInput = [System.Windows.MessageBox]::Show($msg, $title, $Button, $Icon)

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
            $pwd.value = "yourpassword"
            $btSubmit = $doc.getElementById("TANGRAM__PSP_4__submit")
            $btSubmit.click()
        }
        "No" {"Action aborted"; exit}
        
    }
}

$conn.Close()
"Your folder is good."

  



