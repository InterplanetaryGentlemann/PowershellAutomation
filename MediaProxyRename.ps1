#Seth Wagner
#Media Proxy Problem
#04/02/2022

#Import necessary modules for the code to function

Import-Module Microsoft.Powershell.Management

#main - Body of the script, calls all of the other functions of the script in the proper order. Open file dialog and get the user's selection, then import all the file paths in that folder and rename each entry 

function main{

$Files = SelectFolder 

RenameFiles -Files $Files

exit
}

#SelectFolder opens a file browsing dialog where a user can select a folder, the folder is stored as the full file path and returned as the $folder variable

function SelectFolder{

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    
    $foldername.Description = "Select a folder"
    
    $foldername.rootfolder = "Desktop"
    
    $foldername.SelectedPath = $initialDirectory

    #If 

    if($foldername.ShowDialog() -eq "OK")
    
    {
    
        $folder += $foldername.SelectedPath
    
    }

    return $folder

}

#RenameFiles takes each file located in the selected folder and renames it.

function RenameFiles{

    param($Files)

        #An incrementing number, starting from zero that will be the end of the filename

        $increment = 0

        #Separates just the folder name from the full path so it can be reused as the first part of the filename

        $FolderName = $Files | Split-Path -leaf

        # Gets the files stored in the filepath and sorts them by name

        $Rename = @(Get-ChildItem -Attributes !Directory+!System -File $Files | Sort-Object) 

        #Foreach file in the $Rename variable, rename the file with the folder name_Incrementing number and the same file extension 

        foreach($File in $Rename){

        #($Files + "\" + $File) Appends the filename to the end of the path given for the parent folder

        Rename-Item -Path ($Files + "\" + $File) -NewName ($FolderName + ('_{0:D3}' -f $increment++) + $File.Extension)

        }

        exit
}


main