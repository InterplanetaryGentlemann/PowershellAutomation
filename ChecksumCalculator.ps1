#Seth Wagner 
#File Checksum Verification
#11/05/2019

#Run this command before running this script: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted

#Import neccessary modules

Import-Module Microsoft.Powershell.Management

#main - Primary runtime. Calls other functions

function main{

    GetFile

    UserInput

    VerifyHash -File $File -Algorithm $Algorithm -Vhash $Vhash

}

#GetFile - Opens a file Dialog where the user selects the file to generate a hash for

function GetFile{

param($File)

        #Opens a file explorer window that only allows for .csv files

        $openFileDialog = New-Object windows.forms.openfiledialog   
           
            $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
           
            $openFileDialog.title = "Select File to Verify"   
           
            $openFileDialog.filter = "File (*.*)|*.*"    
           
            $openFileDialog.ShowHelp = $True              
            
            $result = $openFileDialog.ShowDialog() 
    
            $result 
    
        #When the user selects a file, checks to ensure it exists and imports it into powershells runtime. 
       
        if($result -eq "OK"){    
                       
            $OpenFileDialog.filename   
           
            $OpenFileDialog.CheckFileExists 
             
            $script:File = $OpenFileDialog.filename  
             
            Write-Host "File Imported!" -ForegroundColor Green
             
        } 
        
        else{ 
        
        Write-Host "File Import Cancelled!" -ForegroundColor Red

        exit
        
        }

}

#UserInput - GUI window that has the user select a hashing algorithm and input the hash to verify against.

function UserInput{
    
    param($Algorithm, $Vhash)

        #Adds the neccessary assembly for th gui window

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        #Sets up the initial form screen

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Checksum Calculator'
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

        #Label for the text box that asks for the Hash

        $hashlabel = New-Object System.Windows.Forms.Label
        $hashlabel.Location = New-Object System.Drawing.Point(10,20)
        $hashlabel.Size = New-Object System.Drawing.Size(280,20)
        $hashlabel.Text = 'Verification Hash:'
        $form.Controls.Add($hashlabel)

        #Text Box for the user's input

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,40)
        $textBox.Size = New-Object System.Drawing.Size(260,20)
        $form.Controls.Add($textBox)

        #Label for the option menu

        $methodlabel = New-Object System.Windows.Forms.Label
        $methodlabel.Location = New-Object System.Drawing.Point(10,75)
        $methodlabel.Size = New-Object System.Drawing.Size(280,20)
        $methodlabel.Text = 'Hash Algorithm Used:'
        $form.Controls.Add($methodlabel)

        #Option menu where user selects which criteria of printers should be removed

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10,100)
        $listBox.Size = New-Object System.Drawing.Size(260,20)
        $listBox.Height = 40

        #Option menu selections

        [void] $listBox.Items.Add('SHA1')
        [void] $listBox.Items.Add('SHA256')
        [void] $listBox.Items.Add('SHA384')
        [void] $listBox.Items.Add('SHA512')
        [void] $listBox.Items.Add('MD5')

        #Options for showing the 

        $form.Controls.Add($listBox)
        $form.Add_Shown({$textBox.Select()})
        $form.Topmost = $true

        #Collects which button (OK, Cancel) the user chooses.

        $result = $form.ShowDialog()

        #If user clicks OK, set variables to the provided inputs, else, generate an error

        if ($result -eq [System.Windows.Forms.DialogResult]::OK){
            
            if([String]$textbox.Text -ne "" -and [String]$listbox.SelectedItem -ne ""){
                    
                    $script:Vhash = $textBox.Text

                    $script:Algorithm = $listBox.SelectedItem

            }


            else{

            Write-Error 'Please enter the Original hash and the Algorithm to use!'

            exit
            
            }

        }

        #If user selects cancel, terminates script

        elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel){

        exit

        }
   
    return $Algorithm, $Vhash

}

#VerifyHash - Generates a hash from the selected file and runs a comparison to the given hash.

function VerifyHash{

    param($File, $Algorithm, $Vhash)

    $Fhash = Get-FileHash -Algorithm $Algorithm -Path $File
    
    if($Fhash.Hash -eq $Vhash){

        Write-Host "Checksum Matches! Original: $Vhash Generated: "$Fhash.Hash"" -ForegroundColor Green

    }

    elseif($Fhash.Hash -ne $Vhash){

        Write-Host "Checksums do not match! Original: $Vhash Generated: "$Fhash.Hash"" -ForegroundColor DarkRed

    }

}

main