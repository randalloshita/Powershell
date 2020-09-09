#Add AD User, mailbox and
#Fill out Employee Checklist form

Write-host "Welcome. Create AD users here. This module only creates new accounts based on a template. Needs mailbox creation. Hit [ENTER] to continue."


$Credential = Get-Credential

#Initials
Write-host "Enter your initials for checkmarking:"
$Initials = Read-Host

#Get files info
$EmployeeChecklistFolder = "P:\EMPLOYEE CHECKLIST\Z_Pending"

Write-Host "Enter the name of the file."
$EmployeeChecklistFileName = Read-Host
$EmployeeChecklistPath = join-path $EmployeeChecklistFolder $EmployeeChecklistFileName

#Read Employeechecklist.xlxs for name of new user
#Convert the Excel into a Powershell object. Requires Import-Excel module.
Install-Module ImportExcel -Scope CurrentUser
#$Excel = Open-ExcelPackage -Path "P:\EMPLOYEE CHECKLIST\Z_Pending\New Employee Checklist_2020-9-1_TEST, TESTOR_IT.xlsx"
$Excel = Open-ExcelPackage -Path $EmployeeChecklistPath
$Worksheet = $Excel.Workbook.Worksheets["Employee Checklist (For IT)"]
$FullName = $Worksheet.Cells['C4'].Value
$BranchDepartment = $Worksheet.Cells['C5'].Value
$Title = $Worksheet.Cells['C6'].Value
$DisplayName = $FullName + " " + "(" + "$BranchDepartment" + ")"
$Name = $BranchDepartment + " " + "-" + " " + $FullName
$UserPrincipalName = $SamAccountName + "@paccitybank.com"

#Test variables
#write-host "outouts"
#$FullName
#$BranchDepartment
#$Title
#$DisplayName
#$Name

Close-ExcelPackage $Excel

#Creating user fields.
#After reading the name on EmployeeChecklist.xlsx, have admin enter a username.
Write-Host "Enter the SamAccountName for " $FullName
$SamAccountName = Read-Host

$Acknowledge = 'N'

Do
{
    Write-Host "Enter the name of user to COPY from:"
    $UserInstance = Read-Host

    $ADUserInstance = Invoke-Command -ComputerName PCBDC1 -Credential $Credential -Scriptblock {Get-ADUser -filter "name -like '*$USing:UserInstance*'"}

    Write-host "This will be the user AD will COPY from: "
    $ADUserInstance.name

    Write-host "Is this correct? [Y] or [N]"
    $Acknowledge = Read-Host

} While ($Acknowledge -eq 'N')

#New-ADUser 
#Convert Password to secure string
$NewADUserPassword = 'Abcd1234'
$SecureString = ConvertTo-SecureString $NewADUserPassword -AsPlainText -Force

#Create ADUser
Invoke-Command -ComputerName PCBDC1 -Credential $Credential -ScriptBlock {New-ADUser -SamAccountName $SamAccountName -Instance $ADUserInstance.name -Title $Title -Department $BranchDepartment -DisplayName $DisplayName -Name $Name -AccountPassword $SecureString -Description $Title}    

#Wait 5 minutes to sync new user object throughout the domain.
start-sleep -second 300

#Modify Home Directory
$ADUser = Invoke-Command -ComputerName PCBDC1 -Credential $Credential -Scriptblock {Get-ADUser -filter "name -like '*$USing:SamAccountName*'" -properties *}
$ADUserHomeDirectory = $ADUser.HomeDirectory

#Create Home Directory and apply permissions
New-Item -Path $ADUserHomeDirectory -Type directory -Credential $Credential
New-SmbShare -Name $SamAccountName -Path $ADUserHomeDirectory -FullAccess "Everyone"
#perissions
Set-ACL -Path $ADUserHomeDirectory 

#Create mailbox
Write-host "Do you want to create a mailbox? [Y] Yes, [N] No."
$CreateMailbox = Read-Host

If ($CreateMailbox -eq 'Y') 
{
    #New mailbox
    Invoke-Command -ComputerName PCBEXCHANGE -Credential $Credential -ScriptBlock {New-Mailbox -Name $FullName -UserPrincipalName $UserPrincipalName -Alias $UserPrincipalName} 
    write-host "Mailbox created."
}
Else
{
    write-host "No mailbox needed."
}


#Fill out Employee Checklist form
#$Excel = Open-ExcelPackage -Path "P:\EMPLOYEE CHECKLIST\Z_Pending\New Employee Checklist_2020-9-1_TEST, TESTOR_IT.xlsx"
$Excel = Open-ExcelPackage -Path $EmployeeChecklistPath
$Worksheet = $Excel.Workbook.Worksheets["Employee Checklist (For IT)"]
$Worksheet.Cells['C11'].Value = $SamAccountName 
$Worksheet.Cells['C14'].Value = $SameAccountName + '@paccitybank.com'
$Worksheet.Cells['C15'].Value = $ADUserHomeDirectory
$Worksheet.Cells['J11'].Value = $Initials
$Worksheet.Cells['J12'].Value = $Initials
$Worksheet.Cells['J14'].Value = $Initials
$Worksheet.Cells['J15'].Value = $Initials






