Import-Module MSOnline
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
Connect-MsolService
<#Check for MSOnline module  
$Module=Get-Module -Name MSOnline -ListAvailable   
if($Module.count -eq 0)  
{  
 Write-Host MSOnline module is not available  -ForegroundColor yellow   
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No  
 if($Confirm -match "[yY]")  
 {  
  Install-Module MSOnline  
  Import-Module MSOnline 
 }  
 else  
 {  
  Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet.  
  Exit 
 } 
}  
  
#Storing credential in script for scheduling purpose/ Passing credential as parameter   
if(($UserName -ne "") -and ($Password -ne ""))   
{   
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword   
 Connect-MsolService -Credential $credential  
}   
else   
{   
 Connect-MsolService | Out-Null   
}  
#>
 
$Result=""    
$PwdPolicy=@{} 
$Results=@()   
$UserCount=0  
 
#Output file declaration  
$ExportCSV="C:\Users\test\Log\PasswordExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"  

#Loop through each user  
Get-MsolUser -DomainName 'sanofi.com' -All | foreach{  
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.DisplayName 
  $PwdLastChange=$_.LastPasswordChangeTimestamp 
 #$PwdNeverExpire=$_.PasswordNeverExpires 
 #$LicenseStatus=$_.isLicensed 
 #$Print=0 
 

 #Check for password expiry 15 days before
 $date = get-date
 $diff = ($date - $PwdLastChange).Days

 if($diff -eq 75)
 { 
 $Result=@{'Display Name'=$DisplayName;'User Principal Name'=$upn;'Pwd Last Change Date'=$PwdLastChange} 
 $UserCount++
 Write-Output $UserCount
 $Results= New-Object PSObject -Property $Result   
 $Results | Select-Object 'Display Name','User Principal Name','Pwd Last Change Date' | Export-Csv -Path $ExportCSV -Notype -Append  
 }
}
 