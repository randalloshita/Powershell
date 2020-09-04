#Add AD User, mailbox and
#Fill out Employee Checklist form

#$Credential = Get-Credential

#Get files info
#$EmployeeChecklistFolder = "P:\EMPLOYEE CHECKLIST\Z_Pending"

#Write-Host "Enter the name of the file."
#$EmployeeChecklistFileName = Read-Host
#$EmployeeChecklistPath = join-path $EmployeeChecklistFolder $EmployeeChecklistFileName

#Read Employeechecklist.xlxs for name of new user
#Convert the Excel into a Powershell object. Requires Import-Excel module.
#Install-Module ImportExcel -Scope CurrentUser
$Excel = Open-ExcelPackage -Path "P:\EMPLOYEE CHECKLIST\Z_Pending\New Employee Checklist_2020-9-1_TEST, TESTOR_IT.xlsx"
$Worksheet = $Excel.Workbook.Worksheets["Employee Checklist (For IT)"]
$FullName = $Worksheet.Cells['C4'].Value
$FullName

Close-ExcelPackage $Excel

#After reading the name on EmployeeChecklist.xlsx, have admin enter a username.
Write-Host "Enter the username for " $FullName
$Username = Read-Host

Write-Host "Enter the name of user to COPY from. 0 for none."
$UsernameToCopyFrom = Read-Host



#$Name = Invoke-Command -ComputerName PCBDC1 -Credential $Credential -ScriptBlock {Get-ADUser -identity "*Randall Oshita*" | Select Name}
#$Name







