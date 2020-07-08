<#
    .Synopsis
        A PowerShell script that queries Active Directory for all enabled users in the Users OU, 
        sorts them by department and job title, exports the list to csv, then copies the csv to
        the correct environment (test or production).
    
              Name: Get-PhoneList.ps1
            Author: Spencer Alessi
          Modified: 07/08/2020
#>

$obj          = @()
$header       = @()
$employeeList = @()
$dateTime     = Get-Date -UFormat "%D %R"
$departments  = "Executive|Finance|Human Resources|Marketing|Sales|Legal|IT|Customer Service"
$contact      = "$(Get-location)\contact.csv"

$searchQuery = Get-ADUser -Filter * -Properties * -SearchBase "ou=yourusersou,ou=yourdomain,dc=yourdc,dc=com" |
               Where-Object {$_.Enabled -eq $true -and $_.department -ne $NULL} |
               Select-Object department, name, title, mail, telephonenumber

$sortedEmployees = $searchQuery | 
                 Sort-Object {$departments.IndexOf($_.department)}, 
                 @{expression={$_.Title -match "(President)|(Controller)"}; descending=$true}, 
                 @{expression={$_.Title -match "(VP)"}; descending=$true}, 
                 @{expression={$_.Title -match "(Director)"}; descending=$true},
                 @{expression={$_.Title -match "(Manager)"}; descending=$true}, 
                 @{expression={$_.Title -match "(Supervisor)"}; descending=$true}, Name

foreach ($employee in $sortedEmployees) {
    if ($employee.telephonenumber){
        $employeeList = [ORDERED]@{
            department = $employee.department
            name       = $employee.name
            title      = $employee.title
            email      = $employee.mail.ToLower()
            extension  = $employee.telephoneNumber.Split('x')[1]
        }             
        $obj += New-Object PSObject -Property $employeeList
    } else {
        #no extension, substitute N/A
        $employeeList = [ORDERED]@{
            department = $employee.department
            name       = $employee.name
            title      = $employee.title
            email      = $employee.mail.ToLower()
            extension  = "N/A"
        }             
        $obj += New-Object PSObject -Property $employeeList
    }
}

$header = @"
###############################################################################
#                      Active Directory User Export                           #
###############################################################################
# Created By:                                                                 #
#    $(Get-location)\$($MyInvocation.MyCommand.Name)                          #
# Fields:                                                                     #
#    Department, Full Name, Title, Email, Telephone Extension                 #
# Sorted by:                                                                  #
#    Department                                                               #
# Updated:                                                                    #
#    $($dateTime)                                                             #
###############################################################################
"@

$header | Set-Content $contact

$obj | ConvertTo-Csv -NoTypeInformation -Delimiter "," |
       ForEach-Object {$_ -replace '"',''}  |
       Select-Object -Skip 1 |
       Add-Content $contact

# copy contact.csv to a folder based on environment
if ($env:COMPUTERNAME -like "tst*"){
    # copy to test
    Copy-Item $contact -Destination "\\tstserver\scripts\phonelist"
} elseif ($env:COMPUTERNAME -like "prod*") {
    # copy to production
    Copy-Item $contact -Destination "\\prodserver\scripts\phonelist"
} else {
    Send-MailMessage -To youremail@domain.com `
                     -From issues@domain.com `
                     -SmtpServer mail.domain.com `
                     -Subject "Get-PhoneList contact.csv copy failed" `
                     -Body "The contact.csv file from Get-PhoneList.ps1 failed to copy to the server."
}
