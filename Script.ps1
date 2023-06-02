# Get the directory of the current script
$scriptDir = $PSScriptRoot

# Find the first CSV file in the script directory
$csvPath = Get-ChildItem -Path $scriptDir -Filter "*.csv" | Select-Object -First 1 -ExpandProperty FullName

# Check if a CSV file was found
if ($null -eq $csvPath) {
    Write-Host "No CSV file found in the script directory: $scriptDir"
    exit
}

# Import the user data from the CSV file
$userData = Import-Csv $csvPath -Delimiter ","

# Define the name of the directory to create for the log files
$logDirName = "LogFiles"

# Combine the script directory with the log directory name to get the full path
$logDir = Join-Path -Path $scriptDir -ChildPath $logDirName

# Create the log directory if it doesn't already exist
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory
}

# Define the log file base name
$logBaseName = "ADUserCreationErrors"

# Check for existing log files and create a new one with an incremented index
$logIndex = 1
while (Test-Path -Path "$logDir\$logBaseName$logIndex.log") {
    $logIndex++
}

# Define the log file path
$logPath = "$logDir\$logBaseName$logIndex.log"

# Get the current domain
$domain = Get-ADDomain
# Extract the domain name from the domain's DNSRoot
$domainName = $domain.DNSRoot

# Loop through the user data and create new users in Active Directory
foreach ($user in $userData) {
    # Set the user properties
    $username = $user.Username
    $fullName = $user.FullName
    $firstName, $lastName = $fullName -split ' ', 2
    $email = $user.Email
    # Split the domain name into DCs and join them with ',DC='
    $dc = 'DC=' + ($domainName -split '\.' -join ',DC=')
    # Set the OU and DC path
    $ou = "OU=$($user.OU),$dc"
    
    # Check if the user has a password defined
    if ($user.Password) {
        # Convert the password into a secure string that Active Directory accepts
        $securePassword = ConvertTo-SecureString $user.Password -AsPlainText -Force
    }

    # Create the new user in Active Directory
    $newUser = $null
    try {
        $newUser = New-ADUser `
            -Name "$firstName $lastName" `
            -SamAccountName $username `
            -UserPrincipalName "$username@$domainName" `
            -AccountPassword $securePassword `
            -GivenName $firstName `
            -Surname $lastName `
            -EmailAddress $email `
            -Enabled $true `
            -Path $ou `
            -PassThru
    }
    catch {
        # Log the error message if the user creation fails
        $errorMessage = @{
            Timestamp = Get-Date -Format o
            Level     = 'ERROR'
            Message   = "Error creating user: $username"
            Details   = @{
                FullName  = $fullName
                FirstName = $firstName
                LastName  = $lastName
                Email     = $email
                OU        = $ou
                Error     = $_.Exception.Message
            }
        }
        Write-Host ($errorMessage | ConvertTo-Json)
        Add-Content -Path $logPath -Value ($errorMessage | ConvertTo-Json)
    }

    # Check if the user was created
    if ($null -ne $newUser) {
        # Log the success message
        $successMessage = @{
            Timestamp = Get-Date -Format o
            Level     = 'INFO'
            Message   = "Successfully created user: $username"
            Details   = @{
                FullName  = $fullName
                FirstName = $firstName
                LastName  = $lastName
                Email     = $email
                OU        = $ou
            }
        }
        Write-Host ($successMessage | ConvertTo-Json)
        Add-Content -Path $logPath -Value ($successMessage | ConvertTo-Json)
    }

    # Add the new user to the Remote Desktop Users group
    try {
        $groupName = "Remote Desktop Users"
        Add-ADGroupMember -Identity $groupName -Members $username
        $successMessage = @{
            Timestamp = Get-Date -Format o
            Level     = 'INFO'
            Message   = "Successfully added user to group: $groupName"
            Details   = @{
                FullName  = $fullName
                FirstName = $firstName
                LastName  = $lastName
                Email     = $email
                OU        = $ou
                Group     = $groupName
            }
        }
        Write-Host ($successMessage | ConvertTo-Json)
        Add-Content -Path $logPath -Value ($successMessage | ConvertTo-Json)

        # Log the error message if the user creation fails
    }
    catch {
        $errorMessage = @{
            Timestamp = Get-Date -Format o
            Level     = 'ERROR'
            Message   = "Error adding user to group: $groupName"
            Details   = @{
                FullName  = $fullName
                FirstName = $firstName
                LastName  = $lastName
                Email     = $email
                OU        = $ou
                Group     = $groupName
                Error     = $_.Exception.Message
            }
        }
        Write-Host ($errorMessage | ConvertTo-Json)
        Add-Content -Path $logPath -Value ($errorMessage | ConvertTo-Json)
        continue
    }

}