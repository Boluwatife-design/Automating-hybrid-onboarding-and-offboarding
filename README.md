# Hybrid Employee onboarding-and-offboarding Automation (PowerShell)

A PowerShell script for automating new-hire account provisioning across on-premises Active Directory, Microsoft Entra ID, and Exchange Online.

# Onboarding
## Overview
The script runs preflight checks before creating any objects, provisions the Active Directory user account, triggers an Entra Connect delta sync, assigns a Microsoft 365 usage location and license, adds group and distribution list memberships, grants shared mailbox permissions, and prints a one-time summary with the temporary password without writing secrets to the log file.

## Key Features
- Preflight validation: Verifies the target Organizational Unit, manager account, naming collisions, Microsoft 365 groups, distribution lists, and shared mailboxes before provisioning begins.

- Hybrid sync & licensing: Creates the on-premises AD user, triggers an Entra Connect delta sync, waits for the user in Entra ID, sets the usage location, and assigns a Microsoft 365 license.

- Access & mailbox setup: Adds Microsoft 365 group memberships via Microsoft Graph and configures distribution lists plus Full Access and Send As permissions on shared mailboxes in Exchange Online once the mailbox is ready.

- Safe logging & alerts: Writes operational logs without storing the generated temporary password, and optionally posts status notifications to a Microsoft Teams webhook.
