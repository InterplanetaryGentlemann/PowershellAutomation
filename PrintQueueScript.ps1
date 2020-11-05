#Seth Wagner - Computer Services Project Intern
#Print Queue Cleanup Script
#09/03/2019

#Import neccessary modules

Import-Module printmanagement

Import-Module Microsoft.Powershell.Management

#main - Grabs Timestamp of script runtime, calls user input functions and runs the selected Printer Removal Process.

function main{

    #Get timestamp for script runtime, this is used to indicate when the operations in this script ran for accounting purposes.

    $TimeStamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }

    UserInput

    #Pings the server to make sure it is up

    TestConnect -ServerName $ServerName

    #If the server is found, or user overrides, continue running script

    if($ContOP -eq $true){

        #If Printers with less than ten jobs is selected, It will prompt the user for the filepath of the PaperCut CSV report. It will find the printers that match these requirements, run an incremental backup and remove the printers.

        if($WhichPrinters -eq 'Printers with less than ten jobs'){

            PaperCutCSV

            LtPrinters -CSV $CSV -ServerName $ServerName

            IncrementalBackup -ToRemove $ToRemove -TimeStamp $TimeStamp -ServerName $ServerName

            DeleteQueues -Printer $Name -ToRemove $ToRemove -ServerName $ServerName

            CleanPapercut -ServerName $ServerName -TimeStamp $TimeStamp

            Write-Host "Process complete. Printers with less than ten jobs have been removed. The list of printers removed and server backup can be found in C:\PrinterBackups on the target print server." -BackgroundColor Green

            exit
        }

        elseif($WhichPrinters -eq 'Offline Printers'){

            OfflinePrinters -ServerName $ServerName
        
            IncrementalBackup -ToRemove $ToRemove -TimeStamp $TimeStamp -ServerName $ServerName

            DeleteQueues -Printer $Name -ToRemove $ToRemove -ServerName $ServerName

            CleanPapercut -ServerName $ServerName -TimeStamp $TimeStamp

            write-host "Process complete. Offline printers have been removed. The list of printers removed and server backup can be found in C:\PrinterBackups on the target print server." -BackgroundColor Green

            exit
        }

    }



}

#Provides a GUI window that prompts user for the FQDN of the print server to be cleaned and gives options for which printers will be removed

function UserInput{
    param($ServerName,$WhichPrinters)

        #Adds the neccessary assembly for th gui window

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        #Sets up the initial form screen

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Print Server Cleanup'
        $form.Size = New-Object System.Drawing.Size(300,225)
        $form.StartPosition = 'CenterScreen'

        #Adds the Ok Button

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(75,150)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

        #Adds the Cancel Button

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(150,150)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.Text = 'Cancel'
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        #Label for the text box that asks for the FQDN

        $serverlabel = New-Object System.Windows.Forms.Label
        $serverlabel.Location = New-Object System.Drawing.Point(10,20)
        $serverlabel.Size = New-Object System.Drawing.Size(280,20)
        $serverlabel.Text = 'FQDN of the target server:'
        $form.Controls.Add($serverlabel)

        #Text Box for the user's input

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,40)
        $textBox.Size = New-Object System.Drawing.Size(260,20)
        $form.Controls.Add($textBox)

        #Label for the option menu

        $methodlabel = New-Object System.Windows.Forms.Label
        $methodlabel.Location = New-Object System.Drawing.Point(10,75)
        $methodlabel.Size = New-Object System.Drawing.Size(280,20)
        $methodlabel.Text = 'Printers to be removed:'
        $form.Controls.Add($methodlabel)

        #Option menu where user selects which criteria of printers should be removed

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10,100)
        $listBox.Size = New-Object System.Drawing.Size(260,20)
        $listBox.Height = 30

        #Option menu selections

        [void] $listBox.Items.Add('Printers with less than ten jobs')
        [void] $listBox.Items.Add('Offline Printers')

        #Options for showing the 

        $form.Controls.Add($listBox)
        $form.Add_Shown({$textBox.Select()})
        $form.Topmost = $true

        #Collects which button (OK, Cancel) the user chooses.

        $result = $form.ShowDialog()

        #If user clicks OK, set variables to the provided inputs, else, generate an error

        if ($result -eq [System.Windows.Forms.DialogResult]::OK){
            
            if([String]$textbox.Text -ne "" -and [String]$listbox.SelectedItem -ne ""){
                    
                    $script:ServerName = $textBox.Text

                    $script:WhichPrinters = $listBox.SelectedItem

            }


            else{

            Write-Error 'Please enter the FQDN and select the type of printers to be removed!'

            exit
            
            }

        }

        #If user selects cancel, terminates script

        elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel){

        exit

        }
   
    return $ServerName, $WhichPrinters

}

#Tests connection to the Provided print server, allows override if ICMP is not allowed.

function TestConnect{

    param($ServerName)

    $CanReach = Test-Connection -ComputerName $ServerName -Quiet

    if($CanReach -eq $true){

        write-host "Server Found! Continuing operation." -ForegroundColor Green

        $script:ContOP = $true

    }

    else{

        $msgBoxInput =  [System.Windows.MessageBox]::Show('Server Unreachable! Check the FQDN you entered. If the server is up and ICMP is disabled, click OK to continue','Warning!','OKCancel','Warning')

            switch($msgBoxInput){

                'OK'{

                    $script:ContOP = $True

                 }
                
                'Cancel'{

                    write-host "Script exited by user." -ForegroundColor Red -BackgroundColor Black

                    exit

                }

            }
                
    }

}

#Opens the file explorer and prompts for the PaperCut CSV needed for the less then ten jobs function

function PaperCutCSV{

    param($CSV)

        #Opens a file explorer window that only allows for .csv files

        $openFileDialog = New-Object windows.forms.openfiledialog   
           
            $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
           
            $openFileDialog.title = "Select PaperCut CSV File to Import"   
           
            $openFileDialog.filter = "CSV Files (*.csv*)|*.csv*"    
           
            $openFileDialog.ShowHelp = $True   
            

            Write-Host 'This option requires the "Printer usage - Summary" CSV report exported from Papercut on the print server. Please select the CSV file. (see FileOpen Dialog)' -ForegroundColor Green  
    
            
            
            $result = $openFileDialog.ShowDialog() 
    
            $result 
    
        #When the user selects a file, checks to ensure it exists and filters to a new .csv this program can read. 
       
        if($result -eq "OK"){    
                       
            $OpenFileDialog.filename   
           
            $OpenFileDialog.CheckFileExists 

            #Powershell uses the first line of a .csv for the headers, Papercut uses the first two lines for descriptive information. This line takes the content of the papercut csv, omits the first two lines and puts the data from a string variable into a new .csv powershell can recognise.
             
            $script:CSV = Get-Content -Path $openFileDialog.FileName |Select-Object -Skip 2 | Out-String | ConvertFrom-Csv  
             
            Write-Host "PaperCut CSV File Imported!" -ForegroundColor Green
             
        } 
        
        else{ Write-Host "PaperCut CSV File Import Cancelled!" -ForegroundColor Red
        
        }

}

#LtPrinters - filters out which Queues we are ready to delete based on a set of criteria, in this case, 
#we want printers that have had less than ten jobs in a month. This is based off the .csv exported from Papercut

function LtPrinters{
    
    param($CSV,$ServerName)

        #Remove Printers that have less than ten jobs based on the papercut csv.
       
        $script:ToRemove = foreach($Printer in $CSV){
            
            #If amount of jobs is less than ten, Get-Printer from rrcclafp02
    
            if($Printer.Jobs -lt 10){

                Get-Printer -ComputerName $ServerName -Name $Printer."Printer Name"


            }

            
        }

    return $ToRemove

}

function OfflinePrinters{
    
    param($ServerName)
        
        $PrintStatus = Get-Printer -ComputerName $ServerName | select printerstatus, name, computername
        
        $script:ToRemove = foreach($Printer in $PrintStatus){

        
        #If printer is offline, add it to the $ToRemove List

            
            if($Printer.printerstatus -eq 'offline'){

                
                Get-Printer -ComputerName $ServerName -Name $Printer.name


                }

        }
    
    return $ToRemove

}

#IncrementalBackup - Creates a .csv listing the printers removed by this script, can run an incremental backup of the server as well. This part has been commented out due to us already running backups on the print server.

function IncrementalBackup{

    param($ToRemove,$TimeStamp,$ServerName)

        #Checks to see if the filepath C:\PrinterBackups exists to write the backup to. Creates it if it does not exist.
                     
        if(   ( Invoke-Command -ComputerName $ServerName -ScriptBlock {Test-Path -EA Stop C:\PrintBackups})){
            #Export a list of the deleted printers to a .csv file in the PrinterBackups folder. Timestamp will be included on the filename.

            $ToRemove | Invoke-Command -ComputerName $ServerName -ScriptBlock {Export-Csv -Path C:\PrinterBackups\RemovedPrinters$using:TimeStamp.csv}

            #Run a Backup of the File Server. Timestamp will be included in the filename of the backup file. !This will only work if the account running the script has the neccessary permissions!

            Invoke-Command -ComputerName $ServerName -ScriptBlock {C:\Windows\System32\spool\tools\PrintBrm.exe -B -S \\$using:ServerName -F C:\PrinterBackups\printserverbackup$using:TimeStamp.printerExport} 
            
            Write-Host "Backup saved to C:\PrinterBackups on the Print Server" -ForegroundColor Green

        }

         
        
        else { 
        
        Invoke-Command -ComputerName $ServerName -ScriptBlock {New-Item -ItemType directory -Path C:\PrinterBackups}

        $ToRemove | Invoke-Command -ComputerName $ServerName -ScriptBlock {Export-Csv -Path C:\PrinterBackups\RemovedPrinters$using:TimeStamp.csv}

        Invoke-Command -ComputerName $ServerName -ScriptBlock {C:\Windows\System32\spool\tools\PrintBrm.exe -B -S \\$using:ServerName -F C:\PrinterBackups\printserverbackup$using:TimeStamp.printerExport}

        Write-Host "Backup saved to C:\PrinterBackups on the Print Server" -ForegroundColor Green
        
        }
       

    return

}

#DeleteQueues - Deletes print queues based on the filtered output from main{}.

function DeleteQueues{
   
    param($Printer,$ToRemove,$ServerName)

        #For Each Loop that deletes printers from rrcclafp02

        foreach($Printer in $ToRemove){

            #Remove Printer from rrcclafp02

            Remove-Printer -ComputerName $ServerName -Name $Printer.name

        }

    return

}

#Function that checks with papercut and the print server and removes printers that no longer exist from papercut.

function CleanPapercut{

    param($ServerName, $TimeStamp)
       
        $Papercut = Invoke-Command -ComputerName $ServerName -ScriptBlock{ & "C:\Program Files\PaperCut NG\server\bin\win\server-command " list-printers } | Out-String 

        $PapercutPrinters = $Papercut -replace ".*\\" | ConvertFrom-csv

        $PrintServer = Get-Printer -ComputerName $ServerName

       
       
        #Creates a new list that contains printers from $PapercutPrinters that do not exist in $PrintServer
        
        $NotExist = $PapercutPrinters | ? {$PrintServer.Name -notcontains $_."!!templateprinter!!"}

        
        
        #Exports a copy of the list of removed Papercut printers to the backup folder

        $NotExist | Invoke-Command -ComputerName $ServerName -ScriptBlock {Export-Csv -Path C:\PrinterBackups\RemovedPapercut$using:TimeStamp.csv}

            
            
            #Foreach loop removes each printer in the list from papercut
            
            foreach($Printer in $NotExist."!!templateprinter!!"){

                Invoke-Command -ComputerName $ServerName -ScriptBlock{ & "C:\Program Files\PaperCut NG\server\bin\win\server-command " delete-printer $using:ServerName $using:Printer} 

            }
        
}

main