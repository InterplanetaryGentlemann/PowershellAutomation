#Seth Wagner 
#File Checksum Verification
#11/05/2019

#Run this command before running this script: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted

#Import neccessary modules

Import-Module Microsoft.Powershell.Management

function main{

    Get-File

    User-Input

    Verify-Hash -File $File -Algorithm $Algorithm -Hash $Hash
}

function Get-File{

    $Script:File = Read-Host "Please Enter the FilePath"

}

function User-Input{

    $Script:Hash = Read-Host "Please Enter the Original Hash"

    $Script:Algorithm = Read-Host "Pleae Enter the Hash Algorithm"

}

function Verify-Hash{

    param($File, $Hash, $Algorithm)

        $FHash = Get-FileHash -Algorithm $Algorithm -Path $File

        if($Fhash.Hash -eq $Hash){

        Write-Host "Checksum Matches! Original: $Hash Generated: "$Fhash.Hash"" -ForegroundColor Green

    }

    elseif($Fhash.Hash -ne $Vhash){

        Write-Host "Checksums do not match! Original: $Hash Generated: "$Fhash.Hash"" -ForegroundColor DarkRed

    }


}

main