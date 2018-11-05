### Set-Manager Script ###
### Purpose: Set the manager of a large number of accounts - Requires an employee gator link and supervisor's gator link ###
### Requires: An excel spreadsheet mapped to the $managerCSV spreadsheet ###
### Programmed by: Joshua Dargan ###
### Date: 10/19/18 ###
$START_TIME = Get-Date
Write-Host "Running..."   $START_TIME -ForegroundColor Green

$managerCSV = Import-CSV C:\Users\sa-joshuadargan282\Documents\Powershell_Assets\DivisionADTESTING.csv
$managerCSVSize = $managerCSV.Count
$unsuccessfulManagerChangesCSV = "C:\Users\sa-joshuadargan282\Documents\Powershell_Assets\FinalUnsuccessfulManagers.csv"
$successfulManagerChangesCSV = "C:\Users\sa-joshuadargan282\Documents\Powershell_Assets\FinalSuccessfulManagers.csv"


$employees = New-Object System.Collections.Generic.List[System.Object] #List for the employees
$managers = New-Object System.Collections.Generic.List[System.Object] #List for managers (the managers should be at the same index of the employees)

$unsuccessfulEmployees = New-Object System.Collections.Generic.List[System.Object] #List for the employees
$unsuccessfulManagers = New-Object System.Collections.Generic.List[System.Object] #List for the employees

#Reads in the $managerCSV to the $employees and $managers lists
function Set-EmployeeManagerLists
{
    Param()
    for ($i = 0; $i -lt $managerCSVSize; $i++)
    {
        if ($managerCSV.EmpGL[$i].Length -gt 0 -AND $managerCSV.SupGL[$i].Length -gt 0)
        {
            $employees.Add($managerCSV.EmpGL[$i])
            $managers.Add($managerCSV.SupGL[$i])
        }
        else 
        {
            $unsuccessfulEmployees.Add($managerCSV.EmpGL[$i])
            $unsuccessfulManagers.Add($managerCSV.SupGL[$i])
        }

    }
}

#Sorts the lists alphabetically according to the employee's gatorlinks to check for duplicates
function Sort-EmpSupLists
{
    Param()
    $empSize = $employees.Count
    if ($empSize -eq $managers.Count -AND $empSize -gt 0)
    {
        #Bubble sorts the lists of strings
        for ($i = 0; $i -lt $empSize; $i++)
        {
            for ($j = 0; $j -lt ($empSize - 1); $j++)
            {
                if ($employees[$j] -gt $employees[$j + 1])
                {
                    #If out of order, swaps the employees and managers.
                    $temp = $employees[$j]
                    $employees[$j] = $employees[$j + 1]
                    $employees[$j + 1] = $temp

                    $temp = $managers[$j]
                    $managers[$j] = $managers[$j + 1]
                    $managers[$j + 1] = $temp
                }
            }
        }
    }
}

#Removes the duplicates
function Remove-DuplicatesInEmpList
{
    Param()
    $count = 0
    while ($count -lt ($employees.Count - 1))
    {
        #If two elements are equal
        if ($employees[$count] -eq $employees[$count + 1]) 
        {
            #Continues to delete the value until there are no more repeats
            $repeatedVal = $employees[$count]
            $continuedRepeats = $true
            while ($continuedRepeats -AND $count -lt $employees.Count)
            {
                if ($employees[$count] -eq $repeatedVal) #Checks if they're equal
                {
                    #Adds them to the unsuccessful changes
                    $unsuccessfulEmployees.Add($employees[$count])
                    $unsuccessfulManagers.Add($managers[$count])
                    #Removes them from the list
                    $employees.RemoveAt($count)
                    $managers.RemoveAt($count)
                }
                else
                {
                    $continuedRepeats = $false
                    $count--
                }
            }
            
        }
        $count++
    }
}

function Edit-ForStudentEmp
{
    for ($i = 0; $i -lt $employees.Count; $i++)
    {
        try
        {
            $name = "sa-" + $employees[$i] 
            $checkForError = Get-ADUser $name
            $employees[$i] = $name
        }
        catch
        {
            #Does nothing when there is an error (This is a non-student employee)
        }
    }
}

#Sets the manager for each employee in the list
function Set-Manager
{
    Param()
    for ($i = 0; $i -lt $managers.Count -AND $i -lt $employees.Count; $i++)
    {
        try
        {
            Get-ADUser $employees[$i] | Set-ADUser -Manager $managers[$i]
            #FIXME https://www.faqforge.com/powershell/append-data-text-file-using-powershell/ $resultOfManagerChanges
        }
        catch
        {
            $unsuccessfulEmployees.Add($employees[$i])
            $unsuccessfulManagers.Add($managers[$i])
            #Write-Host Error with $unsuccessfulEmployees[$i] -ForegroundColor Red
        }
    }
}

#Writes to the $unsuccessfulManagerCSV file
function Output-UnsuccessfulManagerChanges
{
    Param()
    $unsuccessfulChangeTable = New-Object System.Data.DataTable "UnsuccessfulChanges"
    $empGLColumn = New-Object System.Data.DataColumn "EmpGL",([string])
    $supGLColumn = New-Object System.Data.DataColumn "SupGL",([string]) 

    $unsuccessfulChangeTable.Columns.Add($empGLColumn)
    $unsuccessfulChangeTable.Columns.Add($supGLColumn)

    for ($i = 0; $i -lt $employees.Count; $i++) #$unsuccessfulEmployees.Count
    {
        $row =$unsuccessfulChangeTable.NewRow()
        $row."EmpGL" = $unsuccessfulEmployees[$i]
        $row."SupGL" = $unsuccessfulManagers[$i]
        $unsuccessfulChangeTable.Rows.Add($row)
    }

    $unsuccessfulChangeTable | Export-Csv $unsuccessfulManagerChangesCSV 


}

#Writes to the $successfulManagerCSV file
function Output-SuccessfulManagerChanges
{
    Param()
    $successfulChangeTable = New-Object System.Data.DataTable "SuccessfulChanges"
    $empGLColumn = New-Object System.Data.DataColumn "EmpGL",([string])
    $supGLColumn = New-Object System.Data.DataColumn "SupGL",([string])
    $empLocation = New-Object System.Data.DataColumn "EmpLocation",([string])
    $supLocation = New-Object System.Data.DataColumn "SupLocation",([string]) 

    $successfulChangeTable.Columns.Add($empGLColumn)
    $successfulChangeTable.Columns.Add($supGLColumn)
    $successfulChangeTable.Columns.Add($empLocation)
    $successfulChangeTable.Columns.Add($supLocation)

    for ($i = 0; $i -lt $employees.Count; $i++) #$unsuccessfulEmployees.Count
    {
        $row =$successfulChangeTable.NewRow()
        $row."EmpGL" = $employees[$i]
        $row."SupGL" = $managers[$i]

        #Gets the employees's OU https://superuser.com/questions/1179010/get-only-user-ou-from-active-directory-using-powershell-cli
        $user = Get-ADUser -Identity $employees[$i] -Properties CanonicalName
        $userOU = ($user.DistinguishedName -split ",",2)[1]
        $row."EmpLocation" = $userOU

        #Gets the manager's OU
        $user = Get-ADUser -Identity $managers[$i] -Properties CanonicalName
        $userOU = ($user.DistinguishedName -split ",",2)[1]
        $row."SupLocation" = $userOU

        $successfulChangeTable.Rows.Add($row)
    }

    $successfulChangeTable | Export-Csv $successfulManagerChangesCSV 


}

Set-EmployeeManagerLists
Edit-ForStudentEmp
Sort-EmpSupLists
Remove-DuplicatesInEmpList
Set-Manager
Output-UnsuccessfulManagerChanges
Output-SuccessfulManagerChanges

$END_TIME = Get-Date
Write-Host "Ended."  $END_TIME -ForegroundColor Green