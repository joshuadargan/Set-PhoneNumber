### Joshua Dargan ###
### 10/18/18 ###
### Set Manager Script ###
Write-Host "Running..." -ForegroundColor Green

$serverToChange = "housing.ufl.edu"
#Make sure that the .csv only has User and Number columns
$csv = Import-CSV C:\Users\sa-joshuadargan282\Documents\Powershell_Assets\PhoneNumbers1.csv #CHANGE THIS DEPENDING ON WHO RUNS IT
$csvSize = $csv.Count
$usersToChange = New-Object System.Collections.Generic.List[System.Object]
$phones = New-Object System.Collections.Generic.List[System.Object]
#$cred = Get-Credential #If credentials are needed...


#Formats an array of 7 digit numbers to be in the xxx.xxx.xxxx string format
function Format-PhoneNumbers()
{
    Param ([string[]]$phoneNumbers)
    for ($i = 0; $i -lt $phoneNumbers.Length; $i++) #Iterates through the arrays
    {
        if ($phoneNumbers[$i].Length -eq 7) #Formats a 7 digit to be in the xxx.xxx.xxxx string format
        {
            $formattedNumber = "352." + $phoneNumbers[$i].Substring(0,3) + "." + $phoneNumbers[$i].Substring(3)
            $phoneNumbers[$i] = $formattedNumber
        }
    }
    return $phoneNumbers
}
#Returns an array of the Housing usernames for an array of gatorlinks 
function Get-HousingUserIDFromGatorlink()
{
    Param([string[]]$gatorlinks)
    for ($i = 0; $i -lt $gatorlinks.Length; $i++) #Using the Gatorlinks
    {
        $temp = $gatorlinks[$i]
        $housingUsername = Get-ADUser -Filter {info -eq $temp} -Server $serverToChange
        $gatorlinks[$i] = $housingUsername
    }
    return $gatorlinks
}
#Using the .csv file with User and Number columns of equal length, sets all the corresponding values together in arrays
function Set-ListOfUsersAndPhones
{
    Param()
    for ($i = 0; $i -lt $csvSize; $i++)
    {
        $usersToChange.Add($csv.User[$i])
        $phones.Add($csv.Number[$i])
    }
}


Set-ListOfUsersAndPhones #If you uploaded a .csv list, this sets the arrays to that
#$usersToChange = Get-HousingUserIDFromGatorlink($usersToChange) #Gets Housing usernames
$phones = Format-PhoneNumbers($phones) #Formats to the xxx.xxx.xxxx string format

if ($phones -isnot [System.Array] -AND $usersToChange -isnot [System.Array]) #Checks if it is not an array
{
    Set-ADUser $usersToChange -OfficePhone $phones -Server $serverToChange 
    Write-Host $usersToChange "'s phone # changed to " $phones -ForegroundColor Green
}
elseif ($phones.Count -eq $usersToChange.Count) #Checks if the array lengths are equal
{
    for ($i = 0; $i -lt $usersToChange.Count; $i++)
    {
        Set-ADUser $usersToChange[$i] -OfficePhone $phones[$i] -Server $serverToChange #Changes user's phone number in chosen server
        Write-Host $usersToChange[$i] "'s phone # changed to " $phones[$i] -ForegroundColor Green
    }
}
else #Writes an error message
{
    Write-Host "Error formatting phones! Phone and user list lengths are not equal!" -ForegroundColor Red
}

Write-Host "Finished." -ForegroundColor Green

#Write-Output $phones

#Set-ADUser "chrisdtest" -OfficePhone 555-555-5515








