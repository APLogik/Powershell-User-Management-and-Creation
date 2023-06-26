# User Management and Creation

This script is designed to simplify user management tasks by providing a user-friendly menu-driven interface. It requires an SMTP server for sending email notifications and the Active Directory PowerShell modules to be installed. Additionally, the script needs permission to perform the necessary actions in Active Directory.

## Functionality

The script provides the following options in the main menu:

1. **Create a new account**: This option allows you to create a new user account by providing the required information such as first name, last name, email address, and department. The script will generate a secure password for the account and send an email notification to the administrator with the account details.

2. **Select an existing account**: Use this option to search for existing user accounts based on a search term. The script will display a list of matching accounts and prompt you to select one. Once an account is selected, you can perform various actions on it, such as querying or adding groups, generating and setting a new password, or emailing the account information.

3. **Query or add groups to an account**: This option allows you to query the groups a user account belongs to or add new groups to an existing account. You will be prompted to enter the username of the account you want to modify, and then you can choose to either view the current groups or add new ones.

4. **Generate and set an account's password**: Use this option to generate a new secure password for an existing user account and set it as the account's password. You will need to provide the username of the account you want to modify, and the script will generate a new password and set it accordingly.

5. **Email the account information**: This option allows you to send an email containing the account details of an existing user account. You will be prompted to enter the username of the account, and the script will send an email to the administrator with the account's username, password, full name, email address, department, and description.

## Prerequisites

Before using this script, make sure you have the following:

- An SMTP server for sending email notifications.
- Active Directory PowerShell modules installed.
- Sufficient permissions to perform the required actions in Active Directory.

## Configuration

The script includes the following configuration variables:

- `$script:words`: Path to the CSV file containing a list of words used for generating passwords.
- `$script:OrgUnits`: Array of objects representing organizational units (OU) with their names and paths.
- `$script:EmailSender`: Email address and display name of the sender for account-related notifications.
- `$script:ITOrgName`: Name of the IT organization to be used in the subject line of emails.
- `$script:SMTPServer`: IP address or hostname of the SMTP server to be used for sending emails.
- `$script:EmailBodyTop`: The top part of the email body for account-related notifications.

## Usage

To use this script:

1. Set the required variables in the configuration section of the script.
2. Ensure prerequisites are installed
3. Run the script in a PowerShell environment.
4. Follow the prompts to provide the necessary user information for creating a new account or searching for an existing account.
