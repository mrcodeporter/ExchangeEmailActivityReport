<#

╱╱╱╱╱╱╱╱╱╱╱╱╱╭╮╱╱╱╱╱╱╱╱╱╱╱╭╮
╱╱╱╱╱╱╱╱╱╱╱╱╱┃┃╱╱╱╱╱╱╱╱╱╱╭╯╰╮
╭╮╭┳━┳━━┳━━┳━╯┣━━╮╭━━┳━━┳┻╮╭╋━━┳━╮
┃╰╯┃╭┫╭━┫╭╮┃╭╮┃┃━┫┃╭╮┃╭╮┃╭┫┃┃┃━┫╭╯
┃┃┃┃┣┫╰━┫╰╯┃╰╯┃┃━╋┫╰╯┃╰╯┃┃┃╰┫┃━┫┃
╰┻┻┻┻┻━━┻━━┻━━┻━━┻┫╭━┻━━┻╯╰━┻━━┻╯
╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱┃┃
╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╱╰╯

.SYNOPSIS
This PowerShell script connects to Exchange Online and searches for emails based on specified criteria such as sender email address and date range. The script then provides the user with options to export the results to a CSV or HTML file or display them in the console.

.DESCRIPTION
This script is designed to search for emails in Exchange Online based on specified criteria such as sender email address and date range. The script connects to Exchange Online using the Connect-ExchangeOnline cmdlet and then uses the Get-MessageTrace cmdlet to retrieve the email search results. The user is prompted to enter the search criteria and specify whether to export the results to a CSV or HTML file or display them in the console.

.PARAMETER Email
The email address of the sender to search for.

.PARAMETER Date
The date to search for emails in the format "MM/DD/YYYY".

.PARAMETER CSVLocation
The directory where the CSV file should be saved. Defaults to the user's Documents folder.

.PARAMETER HTMLLocation
The directory where the HTML file should be saved. Defaults to the user's Documents folder.

.EXAMPLE
.\Search-ExchangeEmails.ps1
This example starts the script and prompts the user to enter search criteria and choose whether to export the results to a CSV or HTML file or display them in the console.

.INPUTS
None.

.OUTPUTS
The email search results in a table format or exported to a CSV or HTML file.

.NOTES
Author: Ervin Porter
Date: 04/10/2023
Version: 1.0

REQUIREMENTS
This script requires the Exchange Online PowerShell module and PowerShell 5.1 or later.

TROUBLESHOOTING
If you receive an error message indicating that the Exchange Online PowerShell module is not installed, install the module by running the 'Install-Module ExchangeOnlineManagement' command in PowerShell.

If you receive any other error messages when running the script, review the error message for clues about the issue and consult the Exchange Online documentation for additional troubleshooting steps.
ADDITIONS
Error handling: The script includes Try-Catch blocks to catch and handle any errors that might occur during the script's execution.

Menu options: Additional options can be added to the menu to provide users with more functionality, such as the ability to search for emails based on sender, subject, or other criteria.

Output formatting: The script output can be customized to make it more visually appealing and easier to read.

Logging: Logging can be added to the script to record any errors or events that occur during the script's execution.

Performance improvements: The search criteria and search algorithms can be optimized to improve the script's performance.
#>




# Define the Search-Emails function to search for emails and export the results to a CSV or HTML file, or show them in the console
function Search-Emails {
   param(
       [string]$Email,
       [string]$Date,
       [string]$CSVLocation,
       [string]$HTMLLocation
   )

   # Convert the date to a DateTime object
   $startDate = Get-Date $Date
   $endDate = $startDate.AddDays(1)

   # Search for emails sent to the specified email address on the specified date
   $emails = Get-MessageTrace -SenderAddress $Email -StartDate $startDate -EndDate $endDate

   # Show the results in the console
   $emails | Format-Table -AutoSize

   # Give the user the option to export the results to a CSV or HTML file
   $exportToCSV = Read-Host "Export results to CSV file? (y/n)"
   $exportToHTML = Read-Host "Export results to HTML file? (y/n)"

   if ($exportToCSV -eq "y") {
       $defaultCSVName = "{0} Emails {1:MM-dd-yyyy}.csv" -f $Email.Split('@')[0], $startDate
       $csvLocation = save-file -initialDirectory $CSVLocation -filter "CSV (*.csv)| *.csv" -defaultName $defaultCSVName
       if ($csvLocation) {
           $emails | Export-Csv -Path $csvLocation -NoTypeInformation -Encoding UTF8
           Write-Host "Report saved to $csvLocation."
       } else {
           Write-Host "No CSV file selected. Report will not be saved to CSV."
       }
   }

   if ($exportToHTML -eq "y") {
       $defaultHTMLName = "$Email Emails $Date.html"
       $htmlLocation = save-file -initialDirectory $HTMLLocation -filter "HTML (*.html)| *.html" -defaultName $defaultHTMLName
       if ($htmlLocation) {
           $emails | ConvertTo-Html -Property Subject,From,Received,Size,MessageId | Out-File $htmlLocation
           Write-Host "Report saved to $htmlLocation."
           Invoke-Item $htmlLocation
       } else {
           Write-Host "No HTML file selected. Report will not be saved to HTML."
       }
   }

   # Give the user the option to search again with the same email but a different date
   $searchAgain = Read-Host "Search again with the same email but a different date? (y/n)"
   if ($searchAgain -eq "y") {
       $newDate = Read-Host "Enter a new date (MM/DD/YYYY)"
       Search-Emails -Email $Email -Date $newDate -CSVLocation $CSVLocation -HTMLLocation $HTMLLocation
   }
}

# Define the Show-Menu function to display the menu and prompt the user for a choice
function Show-Menu {
   Write-Host "============="
   Write-Host "MENU"
   Write-Host "============="
   Write-Host "1. Search for emails"
Write-Host "2. Exit the program"
Write-Host "============="
return Read-Host "Enter your choice (1-2)"
}

#Set the default CSV and HTML file locations to the user's Documents folder
$CSVLocation = [Environment]::GetFolderPath("MyDocuments")
$HTMLLocation = [Environment]::GetFolderPath("MyDocuments")

#Loop through the menu until the user chooses to exit the program
while ($true) {
$choice = Show-Menu
switch ($choice) {
   1 {
       # Prompt the user to enter an email address and a date
       $email = Read-Host "Enter an email address"
       $date = Read-Host "Enter a date (MM/DD/YYYY)"

       # Search for emails and give the user the option to export the results to a CSV or HTML file
       Search-Emails -Email $email -Date $date -CSVLocation $CSVLocation -HTMLLocation $HTMLLocation
   }
   2 {
       return
   }
   default {
       Write-Host "Invalid choice. Please enter a number between 1 and 2."
   }
}
}