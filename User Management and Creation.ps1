$script:words = Import-CSV "C:\Scripts\example_dict.csv"
$script:OrgUnits = @(
    [pscustomobject]@{OrgName='ExampleOrg1';OrgPath='OU=ExampleOrg1,OU=Example Users,DC=example,DC=ou,DC=edu'}
    [pscustomobject]@{OrgName='ExampleOrg2';OrgPath='OU=ExampleOrg2,OU=Example Users,DC=example,DC=ou,DC=edu'}
)
$script:EmailSender = "ExampleSupport <examplesupport@example.com>"
$script:ITOrgName = "Example IT Org Name" #to be used in subject line of emails
$script:SMTPServer = "192.168.1.1"
$script:EmailBodyTop = "An IT staff member has created a new user account or set a password and wanted to share the details."


function IsValidEmail { 
    param([string]$EmailAddress)
    try {
        $null = [mailaddress]$EmailAddress
        return $true
    }
    catch {
        return $false
    }
}

Function GeneratePw {
    $word1 = Get-Random -InputObject $words  # Select a random word from the $words array and store it in $word1
    $word2 = Get-Random -InputObject $words  # Select another random word from the $words array and store it in $word2
    $number = Get-Random -Minimum 100 -Maximum 999  # Generate a random three-digit number and store it in $number
    $password = $word1.word + $word2.word + $number  # Combine the two words and the number to form the password
    $password = $password.substring(0,1).toupper() + $password.substring(1).tolower()  # Capitalize the first letter of the password and convert the rest to lowercase
    $SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force  # Convert the password to a secure string
    $PasswordPackage = ,@($password, $SecurePassword)  # Create a package containing both the password as a string and the secure password
    return $PasswordPackage  # Return the password package
}

function CreateNewAccount {
    # Gather First Name
    Write-Host "`nLet's Begin`n`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
    while ($Ready -ne "TRUE") {
        $script:FirstName = Read-Host "First Name?"
        if ($FirstName) {
            $Ready = "TRUE"
        } else {
            Write-Host "Invalid entry, try again"
        }
    }
    $Ready = "FALSE" # resetting for next use

    Write-Host "`n`n"

    # Gather Last Name
    while ($Ready -ne "TRUE") {
        $script:LastName = Read-Host "Last Name?"
        if ($LastName) {
            $Ready = "TRUE"
        } else {
            Write-Host "Invalid entry, try again"
        }
    }
    $Ready = "FALSE" # resetting for next use
    $script:UserName = ($FirstName + "." + $LastName).ToLower()
    $script:FullName = $FirstName + " " + $LastName

    Write-Host "`n`n"

    # Gather Email Address
    while ($Ready -ne "TRUE") {
        $script:UsersEmail = Read-Host "Primary E-Mail Address?"
        if ($UsersEmail) {
            $EmailTest = IsValidEmail $UsersEmail # Returns true if email is valid
            if ($EmailTest) {
                $Ready = "TRUE"
            } else {
                Write-Host "Invalid entry, try again"
                $Ready = "False"
                continue
            }
        } else {
            Write-Host "Invalid entry, try again"
        }
    }
    $Ready = "FALSE" # resetting for next use
    # Now get the current context's email address
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $script:AdminEmail = $searcher.FindOne().Properties.mail

    Clear-Host

    $password = GeneratePw
    $script:PasswordString = $password[0]
    $script:SecurePassword = $password[1]

    Write-Host "`n`nPassword will be set to ~ $script:PasswordString ~ you will be prompted later to email this" -ForegroundColor Yellow -BackgroundColor DarkGreen

    Write-Host "`n`n"

    # Change pw on next logon?
    while ($Ready -ne "TRUE") {
        $script:ChangePWOnLogin = Read-Host "Change PW on Next Login? [y/n]"
        if ($ChangePWOnLogin -eq 'n') {
            $script:ChangePWOnLogin = $False
            $Ready = "TRUE"
        }
        if ($ChangePWOnLogin -eq 'y') {
            $script:ChangePWOnLogin = $True
            $Ready = "TRUE"
        }
    }
    $Ready = "FALSE" # resetting for next use

    Write-Host "`n`n"

    # Setting expiration date
    $script:ExpirationTimer = Read-Host "Days till expiration? 6mo = 180. Press ENTER for never"
    if ([string]::IsNullOrWhiteSpace($ExpirationTimer)) {
        $script:ExpirationTimer = "0"
        Write-Host "Account will not expire" -ForegroundColor Yellow -BackgroundColor DarkGreen
    } else {
        Write-Host "`n"
        $script:ExpirationDate = (Get-Date).AddDays($ExpirationTimer).ToString("MM/dd/yyyy hh:mm tt")
        Write-Host "Account will expire on ****** $ExpirationDate ******" -ForegroundColor Yellow -BackgroundColor DarkGreen
    }

    Clear-Host

    # Selecting the ORG Unit
    # used to declare orgunits here
    Write-Host "`n"
    for ($i = 0; $i -lt $OrgUnits.count; $i++) {
        Write-Host "     $($i):  $($OrgUnits[$i].OrgName)" -ForegroundColor White -BackgroundColor Blue
    }
    Write-Host "`n"

    # Gather Department #
    $script:selectedOrgUnit = $null
    while ([string]::IsNullOrWhiteSpace($script:selectedOrgUnit)) {
        while ($Ready -ne "TRUE") {
            $DepartmentNumber = Read-Host "Department #?"
            if ($DepartmentNumber) {
                $Ready = "TRUE"
            } else {
                Write-Host "Invalid entry, try again"
            }
        }
        $Ready = "FALSE" # resetting for next use
        $script:SelectedOrgUnit = $OrgUnits[$DepartmentNumber].OrgPath
        $script:SelectedOrgUnitNiceName = $OrgUnits[$DepartmentNumber].OrgName
        if ([string]::IsNullOrWhiteSpace($script:SelectedOrgUnit)) {
            Write-Host "Invalid entry, try again"
        }
    }
    Write-Host "You Have Selected ****** $($OrgUnits[$DepartmentNumber].OrgName) ******" -ForegroundColor Yellow -BackgroundColor DarkGreen

    Write-Host "`n`n"

    # Gather Description
    $script:CurrentDate = Get-Date -Format "yyyyMMdd"
    $UserDescription = Read-Host "Brief Description?`nAKA Clues to who this person is or who they work for, what area they are in...etc`n*Created $CurrentDate By $env:UserName* <<< This will be appended REGARDLESS`n`n"
    if ([string]::IsNullOrWhiteSpace($UserDescription)) {
        $script:FullDescription = "Created $CurrentDate By $env:UserName"
    } else {
        $script:FullDescription = "$UserDescription + Created $CurrentDate By $env:UserName"
    }

    # Now we spit out all the config and ask for confirmation before committing
    Clear-Host
    Write-Host "`n`n"
    Write-Host "Before adding groups and emailing people, let's commit these changes:" -ForegroundColor Yellow -BackgroundColor Black
    Write-Host "`n`n"
    Write-Host "Full Name = $FirstName $LastName"
    Write-Host "UserName = $UserName"
    Write-Host "Password = $script:PasswordString"
    Write-Host "`n`n"
    Write-Host "ChangePwOnNextLogin = $ChangePWOnLogin"
    if ($ExpirationTimer -eq "0") {
        Write-Host "Expiration Date = Never"
    } else {
        Write-Host "Expiration Date = $ExpirationDate OR $ExpirationTimer Days"
    }
    Write-Host "E-Mail Address = $UsersEmail"
    Write-Host "Org Unit = $($OrgUnits[$DepartmentNumber].OrgName)"
    Write-Host "Description = $FullDescription"
    Write-Host "`n`n`n"

    while ($Ready -ne "TRUE") {
        $CreateAccount = Read-Host "Are you sure you want to create this account? [y/n]"
        if ($CreateAccount -eq 'y') {
            $Ready = "TRUE"
            Clear-Host
            Write-Host "Creating the account`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
            New-ADUser -Name $FullName -GivenName $FirstName -Surname $LastName -DisplayName $FullName -EmailAddress $UsersEmail -AccountExpirationDate $ExpirationDate -ChangePasswordAtLogon $ChangePWOnLogin -SamAccountName $UserName -UserPrincipalName "$UserName@sattrn.ou.edu" -AccountPassword $SecurePassword -Enabled $true -Path $SelectedOrgUnit -Description $FullDescription -Confirm:$false
            Write-Host "Account Creation Complete!`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
            $AdminEmailBody = "New Account Details:`n`nUsername = SATTRN\$UserName`nPassword = $script:PasswordString`n`nFull Name - $FirstName $LastName`nChangePwOnNextLogin = $ChangePWOnLogin`nExpiration Date = $ExpirationDate`nEMail Address = $UsersEmail`nOrg Unit = $SelectedOrgUnitNiceName`nDescription = $FullDescription"
            $AdminEmailSubject = "UserManagementScript - New Account Created"
            Send-MailMessage -From $EmailSender -To $AdminEmail -Subject $AdminEmailSubject -Body $AdminEmailBody -SMTPServer $SMTPServer
            Write-Host "A copy of this login info was just emailed to YOU, based off the address listed on this context's account in AD`n`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
        }
        if ($CreateAccount -eq 'n') {
            $Ready = "TRUE"
            Write-Host "Account Creation Cancelled"
        }
    }
    $Ready = "FALSE" # resetting for next use
}

function CheckForExistingAccount {
    # Check for existing accounts
    Write-Host "`n`nExisting Account Search Function`n"

    # Set initial values
    $script:UseExistingUser = "TRUE"
    $DoneSearchingForUsers = "FALSE"
    $script:selection = $null

    # Loop until done searching for users
    while ($DoneSearchingForUsers -ne "TRUE") {
        # Prompt for search term
        $searchterm = Read-Host -Prompt "ENTER A SEARCH TERM that would for sure be in their account, first, or last name`n`n"

        # Search for users matching the search term
        $users = Get-ADUser -Filter "anr -like '$searchterm'"

        if ($users.count -gt 1) {
            Write-Host
            Write-Host "Multiple users were found" -ForegroundColor Yellow -BackgroundColor DarkGreen

            # Display found users with their index
            for ($i = 0; $i -lt $users.count; $i++) {
                Write-Host
                Write-Host "$($i): $($users[$i].SamAccountName) | $($users[$i].Name)" -ForegroundColor White -BackgroundColor Blue
                Write-Host
            }

            # Prompt for user selection
            while ([string]::IsNullOrWhiteSpace($script:selection)) {
                # Prompt for user selection or quitting the search
                $UserNumber = Read-Host "`nIf you want to continue with an existing user, enter its number`nPress q to stop searching"

                # Check if the user wants to quit
                if ($UserNumber -eq "q") {
                    $QuitNow = "TRUE"
                }

                # Check if a valid user number is entered
                if ($UserNumber) {
                    $Ready = "TRUE"
                    $script:UseExistingUser = "TRUE"
                }
                else {
                    Write-Host "Invalid entry, try again"
                }
            }

            $Ready = "FALSE" # resetting for next use
            $script:selection = $users[$UserNumber].SamAccountName

            # Validate the user selection
            if ([string]::IsNullOrWhiteSpace($script:selection)) {
                Write-Host "Invalid entry, try again"
            }
            else {
                $DoneSearchingForUsers = "TRUE"
            }

            # Check if the user wants to quit
            if ($QuitNow -eq "TRUE") {
                $script:selection = "quit searching"
                $DoneSearchingForUsers = "TRUE"
            }

            Write-Host "`nYou have selected $selection" -ForegroundColor Yellow -BackgroundColor DarkGreen

            # Proceed with existing user or exit the search
            if ($script:selection -eq "quit searching") {
                $script:UseExistingUser = "FALSE"
                Write-Host "`nExiting Search Function`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
            }
            else {
                $script:UseExistingUser = "TRUE"
                GetAdUserInformation -Username $script:selection
                Write-Host "`nGathering Info on Existing User to use in next steps`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
            }
        }
        else {
            if ([string]::IsNullOrWhiteSpace($users)) {
                Write-Host "`nNothing Found" -ForegroundColor Yellow -BackgroundColor DarkGreen

                # Prompt for stopping the search or continuing
                $StopSearching = Read-Host "Press q to stop searching, press anything else to continue"

                if ($StopSearching -eq "q") {
                    $script:UseExistingUser = "FALSE"
                    break;
                }

                # Continue searching
                continue;
            }
            else {
                Write-Host "`nSingle user found - $($users.SamAccountName)" -ForegroundColor Yellow -BackgroundColor DarkGreen
                $script:selection = $users.SamAccountName

                # Prompt for user confirmation
                while ($Ready -ne "TRUE") {
                    $Ready = Read-Host "`nDo you want to continue with this existing user? [y/n]"

                    if ($Ready -eq "n") {
                        $Ready = "TRUE"
                        $DoneSearchingForUsers = "TRUE"
                        $script:selection = "quit searching"
                    }

                    if ($Ready -eq "y") {
                        $Ready = "TRUE"
                        $DoneSearchingForUsers = "TRUE"
                        $script:UseExistingUser = "TRUE"
                    }
                }

                Write-Host "`nYou have selected $script:selection" -ForegroundColor Yellow -BackgroundColor DarkGreen

                # Proceed with existing user or exit the search
                if ($script:selection -eq "quit searching") {
                    $script:UseExistingUser = "FALSE"
                    Write-Host "`nExiting Search Function`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
                }
                else {
                    $script:UseExistingUser = "TRUE"
                    GetAdUserInformation -Username $script:selection
                    Write-Host "`nGathering Info on Existing User to use in next steps`n" -ForegroundColor Yellow -BackgroundColor DarkGreen
                }
            }
        }
    }
}

function DoEmailing {
    # Now let's e-mail this to the customer and other people
    Write-Host "`n`n"

    # Check if UserName is "None" and get AD user information if needed
    if ($script:UserName -eq "None") {
        CheckForExistingAccount
        GetAdUserInformation -Username $script:UserName
    }

    # Set a default value for the PasswordString if it is empty
    if ([string]::IsNullOrWhiteSpace($script:PasswordString)) {
        $script:PasswordString = "Password was not changed"
    }

    # Prepare an array to store attachments
    $Attachments = @()
    $Attachments += 'C:\Welcome.jpg' # example of adding attachments

    # Ask the technician if Materials Package 1 should be included
    $IncludeMaterialsPackage1 = Read-Host "Include Materials Package 1? [y/n]"
    if ($IncludeMaterialsPackage1 -eq 'y') {
        $MaterialsPackage1Info = "Example text describing MaterialsPackage1."
        $Attachments += 'C:\UserGuide1.docx'
    } else {
        $MaterialsPackage1Info = ""
    }

    # Ask the technician if Materials Package 2 should be included
    $IncludeMaterialsPackage2 = Read-Host "Include Materials Package 2? [y/n]"
    if ($IncludeMaterialsPackage2 -eq 'y') {
        $MaterialsPackage2Info = "Example text describing MaterialsPackage2."
        $Attachments += 'C:\UserGuide2.docx'
    } else {
        $MaterialsPackage2Info = ""
    }

    # Prepare the email subject and body
    $EmailSubject = "$script:ITOrgName - New Account Information"
    $EmailBody = "$EmailBodyTop`n`n
Username = $UserName`n
Password = $script:PasswordString`n`n
$MaterialsPackage1Info`n
$MaterialsPackage2Info"

    # Set the parameters for sending the email
    $SendMailParameters = @{
        From       = $EmailSender
        To         = $UsersEmail
        Subject    = $EmailSubject
        Body       = $EmailBody
        SMTPServer = $SMTPServer
        Attachments = $Attachments
    }

    # Ask if the email should be sent to the customer's primary email
    $EmailCustomer = Read-Host "Send to this User's Primary Email? [y/n]"
    if ($EmailCustomer -eq 'y') {
        # Send the email to the customer
        Send-MailMessage @SendMailParameters
    }

    # Ask if the email should be sent to others
    $EmailOthers = Read-Host "To Others? [y/n]"
    if ($EmailOthers -eq 'y') {
        while ($DoneEmailing -ne "TRUE") {
            Write-Host
            $OtherEmail = Read-Host -Prompt "Enter an email address to send to`n"
            $SendMailParameters.To = $OtherEmail
            # Send the email to the specified email address
            Send-MailMessage @SendMailParameters
            $EmailOthers = Read-Host "Email someone else? [y/n]"
            if ($EmailOthers -eq 'n') {
                $DoneEmailing = "TRUE"
            }
        }
    }
}

function QueryOrAddGroups {
    # Check if UserName is set, otherwise check for existing account
    if ($script:UserName -eq "None") {
        CheckForExistingAccount;
    }

    # If using an existing user, retrieve user information
    if ($UseExistingUser -eq "TRUE") {
        $script:UserName = $selection;
        GetAdUserInformation -UserName $UserName
    }

    # Display current group membership
    write-host "`n`n$UserName is currently a member of the following groups`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
    $Groups = Get-ADPrincipalGroupMembership $UserName | foreach-object {Write-Host $_.name}
    write-host $Groups

    # Prompt to add user to groups
    $AddtoGroup = Read-Host "Do you want to add this user to groups? [y/n]"
    if ($AddtoGroup -eq 'y') {
        clear-host

        # Loop for adding groups
        while ($DoneAddingGroups -ne "TRUE") {
            write-host "`n"

            # Prompt for group search term
            $searchterm = Read-Host -Prompt "Enter a search term for the group you want to add`n"

            # Search for groups matching the search term
            $group = Get-ADGroup -LDAPFilter "(anr=$searchterm)"
            if ($group.count -gt 0) {
                write-host
                Write-Host "Here are the groups that were found:" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

                # Display found groups and prompt for selection
                for ($i = 0; $i -lt $group.count; $i++) {
                    write-host
                    Write-Host "$($i): $($group[$i].SamAccountName)" -Foregroundcolor White -Backgroundcolor Blue
                    write-host
                }
                $selection = Read-Host -Prompt "Input the # of the group to add to this user."
                $SelectedGroup = $group[$selection].SamAccountName

                # Validate selected group and add user to the group
                if ([string]::IsNullOrWhiteSpace($SelectedGroup)) {
                    write-host "Invalid Entry, Restarting";
                    continue;
                }
                write-host "Adding $SelectedGroup to user"
                Add-ADGroupMember -Identity $SelectedGroup -Members $UserName

                # Prompt to add another group
                $AnotherGroup = Read-Host "Group added, add another? [y/n]"
                if ($AnotherGroup -eq 'n') {
                    $DoneAddingGroups = "TRUE"
                }
            } else {
                if ([string]::IsNullOrWhiteSpace($group)) {
                    write-host "Invalid Entry, Restarting";
                    continue;
                } else {
                    write-host
                    write-host "Single Result Returned - $($group.SamAccountName)"
                    write-host

                    # Prompt to add user to the single result group
                    while ($Ready -ne "TRUE") {
                        $AddUserToGroup = Read-Host "Add $FullName to $($group.SamAccountName)? [y/n]"
                        $SelectedGroup = $group.SamAccountName

                        # Add user to the group if confirmed
                        if ($AddUserToGroup -eq 'n') {
                            $Ready = "TRUE"
                        }
                        if ($AddUserToGroup -eq 'y') {
                            Add-ADGroupMember -Identity $SelectedGroup -Members $UserName
                            write-host "Group has been added to user."
                            $Ready = "TRUE"

                            # Prompt to add another group
                            $AnotherGroup = Read-Host "Group added, add another? [y/n]"
                            if ($AnotherGroup -eq 'n') {
                                $DoneAddingGroups = "TRUE"
                            }
                        }
                    }
                    $Ready = "FALSE" # Resetting for next use
                }
            }
        }
    }
}

function SetUserPassword {
	# Check if the UserName variable is set to "None"
	if ($script:UserName -eq "None") {
		CheckForExistingAccount
		GetAdUserInformation -Username $script:UserName
	}

	# Generate a password
	$script:password = GeneratePw
	$script:PasswordString = $password[0]
	Write-Host

	# Retrieve the account expiration date
	$ExpirationDate = (Get-ADUser -Identity $UserName -Properties accountexpirationdate).accountexpirationdate

	# Check if the expiration date is null or empty
	if ([string]::IsNullOrWhiteSpace($ExpirationDate)) {
		$ExpirationDate = "None"
	}

	Write-Host "Account Expiration Date = $ExpirationDate`n" -ForegroundColor Yellow -BackgroundColor DarkGreen

	Write-Host "$UserName's password will be set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen

	# Prompt for password regeneration
	while ($Regen -ne 'n') {
		$Regen = Read-Host -Prompt "Re-Generate the password? [y/n]"
		if ($Regen -eq 'y') {
			$script:password = GeneratePw
			$script:PasswordString = $password[0]
			Write-Host "$UserName's password will be set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen
		}
	}

	Write-Host

	# Prompt for password confirmation
	while ($Ready -ne "TRUE") {
		$SetPassword = Read-Host "Press y to set this password and email it to yourself, n to cancel"
		if ($SetPassword -eq 'y') {
			$Ready = "TRUE"
			clear-host

			# Set the account password
			Set-ADAccountPassword -Identity $UserName -Reset -NewPassword $password[1]

			# Get the current context's email address
			$searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
			$script:AdminEmail = $searcher.FindOne().Properties.mail

			# Compose the email body and subject
			$PWEmailBody = "A password has been set on this account.`n`nUsername = $UserName`nPassword = $PasswordString`n`n"
			$PWEmailSubject = "$script:ITOrgName - Account Password Changed"

			# Send the email with the new password
			Send-MailMessage -From $EmailSender -To $AdminEmail -Subject $PWEmailSubject -Body $PWEmailBody -SMTPServer $SMTPServer

			Write-Host "`n`n Account password has been set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen
		}
		if ($SetPassword -eq 'n') {
			clear-host
			Write-Host "No password was changed" -ForegroundColor Yellow -BackgroundColor DarkGreen
			$Ready = "TRUE"
		}
	}
}

function Menu {
    # Check if UserName is null or empty, set it to "None" if true
    if([string]::IsNullOrWhiteSpace($script:UserName)){
        $script:UserName = "None"
    }

    # Display the main menu
    do {
        Write-Host "`n`n================ Main Menu ================`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
        Write-Host "Currently Selected User: $UserName`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

        # Display menu options
        Write-Host "1: Press '1' to create a new account." -Foregroundcolor White -Backgroundcolor Blue
        Write-Host "2: Press '2' to search and select an account." -Foregroundcolor White -Backgroundcolor Blue
        Write-Host "3: Press '3' to query or add groups to a user." -Foregroundcolor White -Backgroundcolor Blue
        Write-Host "4: Press '4' to set account password." -Foregroundcolor White -Backgroundcolor Blue
        Write-Host "5: Press '5' to email account information." -Foregroundcolor White -Backgroundcolor Blue
        Write-Host "Q: Press 'Q' to exit.`n" -Foregroundcolor White -Backgroundcolor Blue

        # Prompt for user selection
        $menuselection = Read-Host "Please make a selection"

        # Perform actions based on the user selection
        switch ($menuselection){
            '1' {
                # Check for existing account and create a new account if necessary
                CheckForExistingAccount
                if($script:UseExistingUser -eq "FALSE"){
                    CreateNewAccount
                }
            }
            '2' {
                # Check for existing account
                CheckForExistingAccount
            }
            '3' {
                # Query or add groups to a user
                QueryOrAddGroups
            }
            '4' {
                # Set account password
                SetUserPassword
            }
            '5' {
                # Perform emailing
                DoEmailing
            }
        }
    }
    # Repeat the menu until the user selects 'q' to exit
    until ($menuselection -eq 'q')
}

Menu;




