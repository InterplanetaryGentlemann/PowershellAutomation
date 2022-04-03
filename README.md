# PowershellAutomation
Various Process I have automated with Powershell Scripts
# Basic Checksum Calculator
This is a very basic script that I have written for demonstration purposes. It allows for beginners to grasp the basic concepts of getting user input and passing those variable in powershell. It calculates the hash of a given file and then compares it against a given hash
# Checksum Calculator
This is a more robust version of the previous tool that I use personally. Pops up GUI windows for gathering the user input (file, algorithm, hash) and then verifies the calculated file hash with the given one. "Why not use WinMD5?" This tool allows me to calculate more than just MD5 hashes and powershell is much faster at doing so than WinMD5. Primarily I developed this out of frustration of long wait times when using WinMD5 to calculate the hashes of CyberPatriot Images.
# Print Queue Script
This script was made to clean out printers from a Windows Print Server that is utilizing Papercut Print Management based on different criteria. Users can select whether to remove printers based on whether or not they have an Offline status in the Windows Print Server or if the printer has had less than 10 print juobs in it's life time based on a csv report generated through Papercut. The script will comb through the list of printers and find all of the ones that match the selected criteria. It will then create a folder and store a timestamped incremental backup so that these printers can be restored if need be and delete the printers from both the Windows Print Server and the Papercut interface.
# Media Proxy Rename
This script is designed to help out certain film students with the lenghty process of ensurng that all of their video files have unique names
for editing proxies. This script will take all of the files in a folder and it will rename all of the files within it with the folder's name and 
an incrementing 3 digit number.

This allows you to organize your files by whatever criteria (filming days, shot numbers, etc.) and rename them accordingly.
For Example:

You move all of your files for your first shot into a folder named Shot1. Running this script on that folder will take all of 
those files, order them by their current name*, and rename them as Shot1_000, Shot1_001, Shot1_002, etc.

*This means that everything will be ordered alphabetically/numerically when they are fed through the name change 
*If you have mov_005 before mov_002 in your file explorer, the script is going to order it numerically and then make the name change, 
*meaning mov_002 would become Shot1_000 and mov_005 would become Shot1_001. 
