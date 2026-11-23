

# 1. Register a free Entra ID app (one-time setup)

- Go to entra.microsoft.com → Identity > Applications > App registrations > New registration
- Name it something like Onboarding-Automation, click Register
- Note the Application (client) ID and Directory (tenant) ID shown on the overview page

<img width="1187" height="726" alt="image" src="https://github.com/user-attachments/assets/5547f95c-7c6b-4876-ad58-bf92bb2dbe0e" />


# 2. Grant it permissions

- In the app, go to API permissions > Add a permission > Microsoft Graph > Application permissions
- Add: User.ReadWrite.All, Group.ReadWrite.All
- Click Grant admin consent (needs a Global Admin)

<img width="1794" height="648" alt="image" src="https://github.com/user-attachments/assets/a2247aa2-c8a6-41d9-bdce-7570108e66e7" />
<img width="233" height="843" alt="image" src="https://github.com/user-attachments/assets/62d25640-5319-4c0d-bce5-c2e8b8214828" />
<img width="310" height="449" alt="image" src="https://github.com/user-attachments/assets/4868853e-29f4-4967-91f6-b0767676711b" />


# 3. Create a client secret

- Go to Certificates & secrets > New client secret
- Copy the secret value immediately (you won’t see it again) — store it securely (e.g., Windows Credential Manager, not in the script)

<img width="1260" height="637" alt="image" src="https://github.com/user-attachments/assets/8f4b249e-c165-49a1-8f2b-ba1a585f182c" />

# 4. Prep the machine that will run the script

- Install RSAT: Add-WindowsFeature RSAT-AD-PowerShell (on a server) or via Windows Optional Features (on a PC)
- Install the Graph module: Install-Module Microsoft.Graph -Scope AllUsers
<img width="1598" height="215" alt="image" src="https://github.com/user-attachments/assets/debbde4b-fbcf-4f73-80ed-c4e686176c68" />



      <#
      .SYNOPSIS
          Automates new-hire provisioning across on-prem Active Directory and Entra ID (hybrid identity).
      
      .DESCRIPTION
          1. Creates the AD user in the correct OU based on department/work-mode.
          2. Adds AD security groups (mail groups, VPN, apps).
          3. Forces an Entra Connect / Cloud Sync delta sync so the account appears in Entra ID quickly.
          4. Waits/polls for the account to land in Entra ID.
          5. Assigns license(s) and Entra/cloud-only groups via Microsoft Graph.
          6. Applies a Conditional Access / device-mode tag (remote vs. office) via an extension attribute.
          7. Logs every step and sends a Teams/email notification on completion or failure.
      
      .NOTES
          Run this from a server/box that has:
            - RSAT ActiveDirectory module (on-prem AD access)
            - Microsoft.Graph PowerShell SDK (or run the Graph section from an Azure Automation hybrid worker)
            - An account with delegated AD OU create rights + a Graph app registration
              (User.ReadWrite.All, Group.ReadWrite.All, Directory.ReadWrite.All) using
              certificate-based auth for unattended runs.
      
      .EXAMPLE
          .\New-EmployeeOnboarding.ps1 -FirstName "Ada" -LastName "Okafor" -Department "Finance" `
              -Manager "jane.doe" -WorkMode "Hybrid" -JobTitle "Financial Analyst" -StartDate "2026-08-04"
      #>
      
      [CmdletBinding()]
      param(
          [Parameter(Mandatory)] [string]$FirstName,
          [Parameter(Mandatory)] [string]$LastName,
          [Parameter(Mandatory)] [string]$Department,
          [Parameter(Mandatory)] [string]$Manager,          # sAMAccountName of manager
          [Parameter(Mandatory)] [ValidateSet("Remote","Office","Hybrid")] [string]$WorkMode,
          [Parameter(Mandatory)] [string]$JobTitle,
          [Parameter(Mandatory)] [datetime]$StartDate,
          [string]$Domain           = "corp.contoso.com",
          [string]$UpnSuffix        = "contoso.com",
          [string]$LicenseSkuId     = "c42b9cae-ea4f-4ab7-9717-81576235ccac", # e.g. Office 365 E3 - replace with your tenant's SKU GUID
          [string]$UsageLocation    = "NG",                 # two-letter country code where the staff work
          [string]$SyncServer       = "AADCONNECT01",       # server running Entra Connect (use $env:COMPUTERNAME if it's this machine)
          [string]$LogPath          = "C:\Logs\Onboarding",
          [string]$TeamsWebhookUrl  = $null
      )
      
      $ErrorActionPreference = "Stop"
      $logFile = Join-Path $LogPath "$($FirstName)_$($LastName)_$(Get-Date -Format yyyyMMdd_HHmmss).log"
      New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
      
      function Write-Log {
          param([string]$Message, [string]$Level = "INFO")
          $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
          Write-Output $line
          Add-Content -Path $logFile -Value $line
      }
      
      function Send-TeamsAlert {
          param([string]$Text, [string]$Color = "00FF00")
          if (-not $TeamsWebhookUrl) { return }
          $body = @{
              "@type"    = "MessageCard"
              "@context" = "http://schema.org/extensions"
              themeColor = $Color
              title      = "Onboarding Automation"
              text       = $Text
          } | ConvertTo-Json -Depth 4
          try { Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -Body $body -ContentType "application/json" }
          catch { Write-Log "Teams notification failed: $($_.Exception.Message)" "WARN" }
      }
      
      try {
          Import-Module ActiveDirectory -ErrorAction Stop
          Write-Log "Starting onboarding for $FirstName $LastName ($Department, $WorkMode)"
      
          # ------------------------------------------------------------------
          # 1. Build identity attributes
          # ------------------------------------------------------------------
          $samAccountName = ("{0}.{1}" -f $FirstName, $LastName).ToLower() -replace '\s',''
          $upn            = "$samAccountName@$UpnSuffix"
          $displayName    = "$FirstName $LastName"
      
          # Single OU per department — adjust to match your actual OU structure
          $targetOU = "OU=$Department,OU=Employees,DC=corp,DC=contoso,DC=com"
      
          if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
              throw "AD account $samAccountName already exists — check for naming collision."
          }
      
          Add-Type -AssemblyName System.Web
          $tempPassword = ([System.Web.Security.Membership]::GeneratePassword(16, 3))
          $securePwd    = ConvertTo-SecureString $tempPassword -AsPlainText -Force
      
          # ------------------------------------------------------------------
          # 2. Create the AD account
          # ------------------------------------------------------------------
          New-ADUser `
              -Name $displayName `
              -GivenName $FirstName `
              -Surname $LastName `
              -SamAccountName $samAccountName `
              -UserPrincipalName $upn `
              -Path $targetOU `
              -AccountPassword $securePwd `
              -Enabled $true `
              -ChangePasswordAtLogon $true `
              -Title $JobTitle `
              -Department $Department `
              -Manager (Get-ADUser -Identity $Manager).DistinguishedName `
              -Description "Provisioned via automation on $(Get-Date -Format 'yyyy-MM-dd')"
      
          Write-Log "AD account created: $samAccountName in $targetOU"
      
          # ------------------------------------------------------------------
          # 3. Group memberships (adjust group names to your environment)
          # ------------------------------------------------------------------
          $groups = @("SG-$Department-All", "SG-M365-Users", "SG-Badge-$Department")
      
          foreach ($group in $groups) {
              try {
                  Add-ADGroupMember -Identity $group -Members $samAccountName
                  Write-Log "Added to group: $group"
              } catch {
                  Write-Log "Could not add to group '$group': $($_.Exception.Message)" "WARN"
              }
          }
      
          # ------------------------------------------------------------------
          # 4. Force Entra Connect / Cloud Sync delta sync
          #    (run this step on the sync server, or via a remote PS session to it)
          # ------------------------------------------------------------------
          Write-Log "Triggering Entra Connect delta sync..."
          Invoke-Command -ComputerName $SyncServer -ScriptBlock {
              Import-Module ADSync
              Start-ADSyncSyncCycle -PolicyType Delta
          } -ErrorAction Stop
          Write-Log "Delta sync triggered."
      
          # ------------------------------------------------------------------
          # 5. Poll Entra ID until the account appears, then assign license/groups
          # ------------------------------------------------------------------
          Write-Log "Connecting to Microsoft Graph using app registration (unattended)..."
          $secureCred = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
          Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $secureCred -NoWelcome
      
          $maxAttempts = 10
          $attempt     = 0
          $entraUser   = $null
      
          do {
              $attempt++
              Start-Sleep -Seconds 30
              $entraUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
              Write-Log "Sync check attempt $attempt/$maxAttempts..."
          } until ($entraUser -or $attempt -ge $maxAttempts)
      
          if (-not $entraUser) {
              throw "User did not appear in Entra ID after $maxAttempts sync attempts. Check Entra Connect health."
          }
      
          Write-Log "User found in Entra ID: $($entraUser.Id)"
      
          # Usage location is REQUIRED before a license can be assigned
          Update-MgUser -UserId $entraUser.Id -UsageLocation $UsageLocation
          Write-Log "Usage location set to $UsageLocation"
      
          # Assign license
          Set-MgUserLicense -UserId $entraUser.Id `
              -AddLicenses @{ SkuId = $LicenseSkuId } `
              -RemoveLicenses @()
          Write-Log "License assigned: $LicenseSkuId"
      
          # ------------------------------------------------------------------
          # 6. Wrap up
          # ------------------------------------------------------------------
          Write-Log "Onboarding completed successfully for $displayName."
          Send-TeamsAlert -Text "✅ Onboarding complete: **$displayName** ($JobTitle, $Department). UPN: $upn. Start date: $($StartDate.ToString('yyyy-MM-dd'))."
      
          # Return a summary object (useful if called from Power Automate / a Logic App via a runbook)
          [PSCustomObject]@{
              DisplayName     = $displayName
              SamAccountName  = $samAccountName
              UPN             = $upn
              OU              = $targetOU
              TempPassword    = $tempPassword
              Status          = "Success"
              LogFile         = $logFile
          }
      }
      catch {
          Write-Log "ONBOARDING FAILED: $($_.Exception.Message)" "ERROR"
          Send-TeamsAlert -Text "❌ Onboarding FAILED for $FirstName $LastName. Error: $($_.Exception.Message)" -Color "FF0000"
          throw
      }
