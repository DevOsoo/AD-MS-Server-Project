# Name: Naveed Fayyez
# Student ID: 010007666

# Import the Active Directory module
Import-Module ActiveDirectory

# Define the OU name and path
$OUCanonicalName = "Finance"
$OUPath = "OU=$OUCanonicalName,$((Get-ADDomain).DistinguishedName)"
$OUExists = Get-ADOrganizationalUnit -Filter {Name -eq $OUCanonicalName}

# Check if OU exists, then delete it
if ($OUExists) {
    Write-Host "$OUCanonicalName OU already exists. Deleting..."
    Remove-ADOrganizationalUnit -Identity $OUPath -Confirm:$false -Recursive # Add -Recursive here
    Write-Host "$OUCanonicalName OU deleted."
}

# Create the OU
try {
    New-ADOrganizationalUnit -Name $OUCanonicalName -Path ((Get-ADDomain).DistinguishedName) -ProtectedFromAccidentalDeletion $false
    Write-Host "$OUCanonicalName OU created."
}
catch {
    Write-Host "Failed to create $OUCanonicalName OU: $($_.Exception.Message)"
} # Add try-catch block here

# Import the financePersonnel.csv file into the Finance OU with the specified properties
$csvPath = Join-Path $PSScriptRoot "financePersonnel.csv"
$users = Import-Csv $csvPath

foreach ($user in $users) {
    $firstName = $user."First Name"
    $lastName = $user."Last Name"
    $displayName = "$firstName $lastName"
    $postalCode = $user."Postal Code"
    $officePhone = $user."Office Phone"
    $mobilePhone = $user."Mobile Phone"
    $samAccountName = $user."samAccount"

    # Check if the user account name already exists
    if (Get-ADUser -Identity $samAccountName) {
        Write-Host "User account name $samAccountName already exists. Please choose another."
        continue # Skip this user and move on to the next one
    }

    # Create the user account with the specified properties
    $userParams = @{
        GivenName        = $firstName
        Surname          = $lastName
        Name             = $displayName
        SamAccountName   = $samAccountName
        UserPrincipalName = "$samAccountName@$((Get-ADDomain).DNSRoot)"
        DisplayName      = $displayName
        PostalCode       = $postalCode
        MobilePhone      = $mobilePhone
        OfficePhone      = $officePhone
        Path             = $OUPath
        Enabled          = $true
        PasswordNotRequired = $true # Add this parameter here
    }

    try {
        New-ADUser @userParams
    }
    catch {
        Write-Host "Failed to create user account"
    }
}

# Generate an output file for submission
$outputPath = Join-Path $PSScriptRoot "AdResults.txt"
Get-ADUser -Filter * -SearchBase $OUPath -Properties DisplayName, PostalCode, OfficePhone, MobilePhone |
    Select-Object DisplayName, PostalCode, OfficePhone, MobilePhone |
    Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "Output file generated: $outputPath"
