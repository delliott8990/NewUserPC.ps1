$Comments = @' 
Script name: NewUserPC.ps1 
Created on: 5/6/16 
Author: Derrick Elliott
Purpose: Automate Steps for New User On-Boarding
'@ 



## To Do ###############################################################################################
# 1. Change Office Choices in UI prompt (30,31)
# 2. Change Office info in Switch Statement - Office, State Abbreviation, and OU Location (40, 41, & 42)
# 3. Update SERVERNAME name in all paths used with correct F&P name (93,118,130,166,& 187)
########################################################################################################

<#Import AD module#>
import-module activedirectory



#################################################
#   Prompt and populate info for user and pc csvs
#################################################



<#Prompt for Office Choice#>
[int]$choice = read-host -prompt '
-------Choose Office-------
1 = Office One
2 = Office Two
---------------------------
Key in selection and press enter'

<#Assigns correct Office name, state abbreviation, and OU to PC csv#>
Switch ($choice)
{
	1 
	{
        $Office = ""   #Office Name
        $State = "" #State Abbreviation
        $OU = 'OU=Computers,OU=LocationOne,DC=YourDomain,DC=com'
        #OU location must be entered as listed in AD with no spaces after OU= or before comma
    }
    2
    {
        $Office = ""   #Office Name
        $State = "" #State Abbreviation
        $OU = 'OU=Computers,OU=LocationTwo,DC=YourDomain,DC=com'
        #OU location must be entered as listed in AD with no spaces after OU= or before comma
    }
}


<#Prompts and stores info in variables#>
Write-Host "New Colleague Information"
$PCName = Read-Host "PC Name?"
$Firstname = Read-Host "Firstname?"
$Lastname = Read-Host "Lastname?"
$StartDate  = Read-Host "Start Date?"
$Username = Read-Host "Username?"
$Email = $Firstname + '.' + $Lastname + '@sedgwickcms.com'
$Phone = Read-Host "Phone number?"
$Ext = Read-Host "Extension?"
$UserFileName = $Username + '.txt'
$PCFileName = $PCName + '.txt'



######################################
# Create User CSV
######################################



<#Undefined array for storing user info#>
$csvContents = @()

<#Pass info to array for csv#>
$row = New-Object System.Object # Create an object to append to the array
$row | Add-Member -MemberType NoteProperty -Name "Firstname" -Value $Firstname
$row | Add-Member -MemberType NoteProperty -Name "Lastname" -Value $Lastname
$row | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $StartDate
$row | Add-Member -MemberType NoteProperty -Name "Username" -Value $Username
$row | Add-Member -MemberType NoteProperty -Name "Email" -Value $Email
$row | Add-Member -MemberType NoteProperty -Name "Phone" -Value $Phone
$row | Add-Member -MemberType NoteProperty -Name "Extension" -Value $Ext
$csvContents += $row

<#Export array contents to csv file#>
$csvContents | Export-CSV -Path \\SERVERNAME\share\restofpath\$UserFileName



######################################
# Create PC CSV
######################################



<#Undefined array for PC info#>
$newPCcsv = @()

<#Pass info to array for csv#>
$row = New-Object System.Object # Create an object to append to the array
$row | Add-Member -MemberType NoteProperty -Name "PCName" -Value $PCName
$row | Add-Member -MemberType NoteProperty -Name "Office" -Value $Office
$row | Add-Member -MemberType NoteProperty -Name "State" -Value $State
$row | Add-Member -MemberType NoteProperty -Name "Userfirst" -Value $FirstName
$row | Add-Member -MemberType NoteProperty -Name "Userlast" -Value $LastName
$row | Add-Member -MemberType NoteProperty -Name "OU" -Value $OU
$newPCcsv += $row

<#Export array contents to csv#>
$newPCcsv | Export-CSV -Path \\SERVERNAME\share\restofpath\$PCFileName
clear



#################################################
# Update AD Description/OU using PC info csv file
#################################################


#import csv file
$NewPCcsvIn = import-csv \\SERVERNAME\share\restofpath\$PCFileName

<#Import info from csv#>
foreach($test in $NewPCcsvIn)
{
    $pcname = $test.PCName
    $office = $test.Office
    $state = $test.State
    $fname = $test.Userfirst
    $lname = $test.Userlast
    $OU = $test.OU
    
    $NEWDescription = "$office, $state - $lname, $fname"
    
    <#Writes to AD Description#>
    Set-ADComputer $pcname -Description $NEWDescription
    Write-Host 'AD description has been updated'
    
	

    <#AD Search by OU, Move PC to new OU#>
    $TARGETOU = Get-ADOrganizationalUnit -Identity $OU
    Get-ADComputer $pcname | Move-ADObject -TargetPath $TARGETOU.DistinguishedName
    Write-Host 'Moved to ' + $office + ' OU'
}



#################################################
# Generate New Hire Packet using user info csv
#################################################



<#import csv file#>
$NewUserCsvIn = import-csv \\SERVERNAME\share\restofpath\$UserFileName

<#create variables to store csv info#>
foreach($user in $csvContents)
{
    $fn = $user.Firstname
    $ln = $user.Lastname
    $date = $user.StartDate
    $username = $user.Username
    $email = $user.Email
    $phone = $user.Phone
    $ext = $user.Extension
}


<#Create new word object#>
$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $True 

<#Open NewUserTemplate doc#>
#replace servername with F&P name (ex: mspfp02)
$objDoc = $objWord.Documents.Open("\\SERVERNAME\share\restofpath\YourNewHirePacketTemplate.docx") 
$objSelection = $objWord.Selection
$wdReplaceAll = 2 
$wdFindContinue = 1 
 
<#Replace first name#>
$FindText = "<first>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $fn

$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, `
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, `
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll)

<#Replace Last Name#>
$FindText = "<last>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $ln


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 

<#replace username#>
$FindText = "<username>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $username


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 
    
<#replace start date#>
$FindText = "<start>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $date


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 

<#replace email#>
$FindText = "<email>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $email


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 

<#replace phone#>
$FindText = "<phone>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $phone


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 

<#replace ext#>
$FindText = "<ext>"
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$ReplaceWith =  $ext


$a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward, ` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll) 