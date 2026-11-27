

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

# 4. Store the client secret

- Run this in PowerShell as the same account, on the same machine, that will run the script:
  
      New-Item -ItemType Directory -Path C:\Secrets -Force
      Read-Host "Paste client secret" -AsSecureString | Export-Clixml -Path "C:\Secrets\onboarding-secret.xml"
  
The file is encrypted to that account and machine. It cannot be copied to another machine or opened by another account.
- If you no longer have the secret value (Entra only shows it once), create a new one: app registration, then Certificates & secrets, then New client secret, and copy the value immediately.

# 5. Get your license SKU ID

    Connect-MgGraph -Scopes "Organization.Read.All"
    Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId, ConsumedUnits

- O365_BUSINESS_ESSENTIALS = Business Basic
- O365_BUSINESS_PREMIUM = Business Standard

Copy the SkuId for the plan new hires should receive

# 6. Prep the machine that will run the script

- Install RSAT: 

      Add-WindowsFeature RSAT-AD-PowerShell (on a server) or via Windows Optional Features (on a PC)

- Install the Graph module: 

      Install-Module Microsoft.Graph -Scope AllUsers
Close PowerShell and open a fresh admin window after installing the Graph module (see Troubleshooting).
<img width="1598" height="215" alt="image" src="https://github.com/user-attachments/assets/debbde4b-fbcf-4f73-80ed-c4e686176c68" />

# 7. Configure the script
- Save the script as C:\Scripts\New-EmployeeOnboarding.ps1 and set:


                        <#
                        .SYNOPSIS
                            Onboards a new employee in a hybrid environment (on-prem AD + Entra ID + Exchange Online).
                        
                        .DESCRIPTION
                            Preflight : connects to Microsoft Graph (and Exchange Online if needed) and checks that the OU,
                                        the manager and every requested group / list / shared mailbox exist.
                                        If any check fails, NOTHING is created.
                            Step 1    : creates the AD user in OU=<Department> under the Company-Users OU.
                            Step 2    : triggers an Entra Connect delta sync and waits for the user to appear in Entra ID.
                            Step 3    : sets the usage location and assigns the Microsoft 365 license.
                            Step 4    : adds the user to Microsoft 365 groups (Microsoft Graph).
                            Step 5    : waits for the mailbox, then adds the user to distribution lists and grants
                                        shared mailbox access (Exchange Online).
                            Finish    : prints a summary and the temporary password ONCE on screen.
                                        The password is never written to the log file.
                        
                            Groups, lists and shared mailboxes are identified by their EMAIL ADDRESS.
                            Every new hire also gets the defaults set under "DEFAULTS FOR EVERY NEW HIRE" below.
                        
                            Shared mailboxes: the user is granted Full Access (auto-mapped in Outlook) AND Send As.
                        
                        .EXAMPLE
                           $secureSecret = Import-Clixml -Path "C:\Secrets\onboarding-secret.xml"
                        
                            .\New-EmployeeOnboarding.ps1 -FirstName "Jane" -LastName "Doe" -Department "FINANCE" `
                                -Manager "john.smith" -JobTitle "Financial Analyst" -StartDate "2026-12-01" `
                                -M365Groups "finance-team@contoso.com" `
                                -DistributionLists "finance-all@contoso.com" `
                                -SharedMailboxes "accounts@contoso.com","payments@contoso.com" `
                                -ExchangeAdminUpn "admin@contoso.com" `
                                -TenantId "<TENANT-ID>" -ClientId "<CLIENT-ID>" -ClientSecret $secureSecret
                        #>
                        
                        [CmdletBinding()]
                        param(
                            [Parameter(Mandatory)] [string]$FirstName,
                            [Parameter(Mandatory)] [string]$LastName,
                            [Parameter(Mandatory)] [string]$Department,        # must match an OU name under Company-Users exactly
                            [Parameter(Mandatory)] [string]$Manager,           # manager's AD login name (sAMAccountName)
                            [Parameter(Mandatory)] [string]$JobTitle,
                            [Parameter(Mandatory)] [datetime]$StartDate,
                        
                            # Chosen per employee. Use the EMAIL ADDRESS of each item.
                            [string[]]$M365Groups        = @(),
                            [string[]]$DistributionLists = @(),
                            [string[]]$SharedMailboxes   = @(),
                        
                            # Entra app registration (Microsoft Graph, unattended)
                            [Parameter(Mandatory)] [string]$TenantId,
                            [Parameter(Mandatory)] [string]$ClientId,
                            [Parameter(Mandatory)] [securestring]$ClientSecret,
                        
                            # Admin who signs in to Exchange Online (only used when lists or shared mailboxes are requested)
                            [string]$ExchangeAdminUpn,
                        
                            [string]$UpnSuffix       = "contoso.com",
                            [string]$LicenseSkuId    = "<YOUR-LICENSE-SKU-ID>",   # get from Get-MgSubscribedSku
                            [string]$UsageLocation   = "US",
                            [string]$SyncServer      = "<SYNC-SERVER-NAME>",      # server running Entra Connect
                            [string]$LogPath         = "C:\Logs\Onboarding",
                            [string]$TeamsWebhookUrl = $null
                        )
                        
                        $ErrorActionPreference = "Stop"
                        
                        # ---------------- SETTINGS ----------------
                        $BaseOU = "OU=Company-Users,DC=corp,DC=contoso,DC=com"
                        
                        # ---------------- DEFAULTS FOR EVERY NEW HIRE (email addresses) ----------------
                        $DefaultM365Groups        = @()
                        $DefaultDistributionLists = @("employees@contoso.com")   # CHANGE to your real Employees list address
                        $DefaultSharedMailboxes   = @()
                        
                        # ---------------- LOGGING ----------------
                        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
                        $logFile = Join-Path $LogPath ("{0}_{1}_{2}.log" -f $FirstName, $LastName, (Get-Date -Format "yyyyMMdd_HHmmss"))
                        
                        function Write-Log {
                            param([string]$Message, [string]$Level = "INFO")
                            $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
                            $color = switch ($Level) { "ERROR" { "Red" } "WARN" { "Yellow" } default { "Gray" } }
                            Write-Host $line -ForegroundColor $color
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
                            try { Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -Body $body -ContentType "application/json" | Out-Null }
                            catch { Write-Log "Teams notification failed: $($_.Exception.Message)" "WARN" }
                        }
                        
                        function Merge-List {
                            param([string[]]$Defaults, [string[]]$Extra)
                            @(@($Defaults) + @($Extra) | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim().ToLower() } | Select-Object -Unique)
                        }
                        
                        $allM365   = @(Merge-List -Defaults $DefaultM365Groups        -Extra $M365Groups)
                        $allDLs    = @(Merge-List -Defaults $DefaultDistributionLists -Extra $DistributionLists)
                        $allShared = @(Merge-List -Defaults $DefaultSharedMailboxes   -Extra $SharedMailboxes)
                        $needExchange = ($allDLs.Count -gt 0) -or ($allShared.Count -gt 0)
                        
                        $memberships    = New-Object System.Collections.Generic.List[string]
                        $problems       = New-Object System.Collections.Generic.List[string]
                        $adCreated      = $false
                        $samAccountName = $null
                        
                        try {
                            Import-Module ActiveDirectory -ErrorAction Stop
                            Add-Type -AssemblyName System.Web
                  
                      if ($LicenseSkuId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
                          throw "LicenseSkuId is not set. Paste your real SkuId into the script (get it with Get-MgSubscribedSku)."
                      }
                  
                      # Naming convention: first letter of first name + surname, e.g. Jane Doe -> jdoe
                      $samAccountName = ("{0}{1}" -f $FirstName.Trim().Substring(0,1), $LastName.Trim()).ToLower() -replace '\s',''
                      $upn            = "$samAccountName@$UpnSuffix"
                      $displayName    = "$FirstName $LastName"
                      $targetOU       = "OU=$Department,$BaseOU"
                  
                      Write-Log "Starting onboarding for $displayName ($Department) - $upn"
                  
                      # ==================================================================
                      # PREFLIGHT: connect and check everything BEFORE creating anything
                      # ==================================================================
                      Write-Log "Connecting to Microsoft Graph..."
                      $graphCred = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
                      Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $graphCred -NoWelcome
                  
                      if ($needExchange) {
                          Write-Log "Connecting to Exchange Online (a sign-in window will appear)..."
                          Import-Module ExchangeOnlineManagement -ErrorAction Stop
                          if ($ExchangeAdminUpn) { Connect-ExchangeOnline -UserPrincipalName $ExchangeAdminUpn -ShowBanner:$false }
                          else                   { Connect-ExchangeOnline -ShowBanner:$false }
                      }
                  
                      $missing = @()
                  
                      # AD checks
                      try { Get-ADOrganizationalUnit -Identity $targetOU -ErrorAction Stop | Out-Null }
                      catch { $missing += "AD OU not found: $targetOU  (check the -Department spelling)" }
                  
                      $managerDn = $null
                      try { $managerDn = (Get-ADUser -Identity $Manager -ErrorAction Stop).DistinguishedName }
                      catch { $missing += "Manager not found in AD: $Manager  (use the AD login name)" }
                  
                      if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
                          $missing += "AD account $samAccountName already exists - check for a naming collision."
                      }
                  
                      # Microsoft 365 groups
                      $m365Resolved = @()
                      foreach ($addr in $allM365) {
                          $g = Get-MgGroup -Filter "mail eq '$addr'" | Select-Object -First 1
                          if (-not $g) { $missing += "Microsoft 365 group not found: $addr" }
                          elseif ($g.GroupTypes -notcontains "Unified") { $missing += "Not a Microsoft 365 group (if it is a distribution list, pass it in -DistributionLists): $addr" }
                          elseif ($g.GroupTypes -contains "DynamicMembership") { $missing += "Dynamic group - membership is automatic, cannot add manually: $addr" }
                          else { $m365Resolved += $g }
                      }
                  
                      # Distribution lists
                      foreach ($addr in $allDLs) {
                          $dl = $null
                          try { $dl = Get-DistributionGroup -Identity $addr -ErrorAction Stop } catch { $dl = $null }
                          if (-not $dl) { $missing += "Distribution list not found: $addr" }
                          elseif ($dl.IsDirSynced) { $missing += "Distribution list is synced from on-prem AD and cannot be edited in the cloud: $addr" }
                      }
                  
                      # Shared mailboxes
                      foreach ($addr in $allShared) {
                          $sm = $null
                          try { $sm = Get-Mailbox -Identity $addr -RecipientTypeDetails SharedMailbox -ErrorAction Stop } catch { $sm = $null }
                          if (-not $sm) { $missing += "Shared mailbox not found (or it is not a shared mailbox): $addr" }
                      }
                  
                      if ($missing.Count -gt 0) {
                          throw ("Preflight failed. NOTHING was created. Fix these and run again:`n - " + ($missing -join "`n - "))
                      }
                      Write-Log "Preflight passed."
                  
                      # ==================================================================
                      # STEP 1: Create the AD account
                      # ==================================================================
                      $tempPassword = [System.Web.Security.Membership]::GeneratePassword(16, 3)
                      $securePwd    = ConvertTo-SecureString $tempPassword -AsPlainText -Force
                  
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
                          -Manager $managerDn `
                          -Description "Provisioned via automation on $(Get-Date -Format 'yyyy-MM-dd')"
                  
                      $adCreated = $true
                      Write-Log "AD account created: $samAccountName in $targetOU"
                  
                      # ==================================================================
                      # STEP 2: Sync to Entra ID and wait for the user to appear
                      # ==================================================================
                      Write-Log "Triggering Entra Connect delta sync on $SyncServer..."
                      $syncBlock = { Import-Module ADSync; Start-ADSyncSyncCycle -PolicyType Delta | Out-Null }
                      try {
                          if ($SyncServer -ieq $env:COMPUTERNAME -or $SyncServer -ieq "localhost") { & $syncBlock }
                          else { Invoke-Command -ComputerName $SyncServer -ScriptBlock $syncBlock }
                          Write-Log "Delta sync triggered."
                      } catch {
                          Write-Log "Could not trigger the delta sync ($($_.Exception.Message)). Continuing - the normal sync cycle may still pick the user up." "WARN"
                      }
                  
                      $maxAttempts = 20
                      $attempt     = 0
                      $entraUser   = $null
                      do {
                          $attempt++
                          Start-Sleep -Seconds 30
                          $entraUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
                          Write-Log "Waiting for user to appear in Entra ID ($attempt/$maxAttempts)..."
                      } until ($entraUser -or $attempt -ge $maxAttempts)
                  
                      if (-not $entraUser) {
                          throw "User did not appear in Entra ID after $maxAttempts checks. Check Entra Connect health and that the OU is in sync scope."
                      }
                      Write-Log "User found in Entra ID: $($entraUser.Id)"
                  
                      # ==================================================================
                      # STEP 3: Usage location and license
                      # ==================================================================
                      # A brand-new user can take a short while before Graph accepts changes to it, so retry
                      $ulSet = $false
                      for ($u = 1; $u -le 8 -and -not $ulSet; $u++) {
                          try {
                              Update-MgUser -UserId $entraUser.Id -UsageLocation $UsageLocation -ErrorAction Stop
                              $ulSet = $true
                          } catch {
                              Write-Log "Usage location attempt $u/8 not accepted yet: $($_.Exception.Message)" "WARN"
                              if ($u -lt 8) { Start-Sleep -Seconds 15 }
                          }
                      }
                      if (-not $ulSet) {
                          throw "Could not set the usage location after 8 attempts. The AD account and Entra user DO exist - finish the setup by hand."
                      }
                      Write-Log "Usage location set to $UsageLocation. Waiting for it to take effect..."
                  
                      $ulTries = 0
                      do {
                          $ulTries++
                          Start-Sleep -Seconds 10
                          $ulCheck = Get-MgUser -UserId $entraUser.Id -Property UsageLocation -ErrorAction Stop
                      } until ($ulCheck.UsageLocation -eq $UsageLocation -or $ulTries -ge 6)
                  
                      $licensed = $false
                      for ($i = 1; $i -le 6 -and -not $licensed; $i++) {
                          try {
                              Set-MgUserLicense -UserId $entraUser.Id -AddLicenses @{ SkuId = $LicenseSkuId } -RemoveLicenses @() -ErrorAction Stop | Out-Null
                              $licensed = $true
                          } catch {
                              Write-Log "License attempt $i/6 failed: $($_.Exception.Message)" "WARN"
                              if ($i -lt 6) { Start-Sleep -Seconds 15 }
                          }
                      }
                      if (-not $licensed) {
                          throw "The license could not be assigned after 6 attempts. The AD account and Entra user DO exist - assign the license manually, then add the remaining memberships by hand."
                      }
                      Write-Log "License assigned."
                  
                      # ==================================================================
                      # STEP 4: Microsoft 365 groups (Graph)
                      # ==================================================================
                      foreach ($g in $m365Resolved) {
                          try {
                              New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $entraUser.Id -ErrorAction Stop
                              Write-Log "Added to Microsoft 365 group: $($g.Mail)"
                              $memberships.Add("M365 group: $($g.Mail)")
                          } catch {
                              if ($_.Exception.Message -match "already exist") {
                                  Write-Log "Already a member of Microsoft 365 group: $($g.Mail)"
                                  $memberships.Add("M365 group: $($g.Mail) (already member)")
                              } else {
                                  Write-Log "FAILED to add to Microsoft 365 group $($g.Mail): $($_.Exception.Message)" "WARN"
                                  $problems.Add("M365 group: $($g.Mail)")
                              }
                          }
                      }
                  
                      # ==================================================================
                      # STEP 5: Distribution lists and shared mailboxes (Exchange Online)
                      # ==================================================================
                      if ($needExchange) {
                          Write-Log "Waiting for the new mailbox to be created in Exchange Online (usually a few minutes)..."
                          $mbx      = $null
                          $mbxTries = 0
                          do {
                              $mbxTries++
                              try { $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop } catch { $mbx = $null }
                              if (-not $mbx) {
                                  Write-Log "Mailbox not ready yet ($mbxTries/30)..."
                                  Start-Sleep -Seconds 30
                              }
                          } until ($mbx -or $mbxTries -ge 30)
                  
                          if (-not $mbx) {
                              Write-Log "Mailbox was not ready after 15 minutes. Distribution lists and shared mailboxes were NOT applied." "WARN"
                              foreach ($a in $allDLs)    { $problems.Add("Distribution list (pending, mailbox not ready): $a") }
                              foreach ($a in $allShared) { $problems.Add("Shared mailbox (pending, mailbox not ready): $a") }
                          } else {
                              Write-Log "Mailbox is ready."
                  
                              foreach ($addr in $allDLs) {
                                  try {
                                      Add-DistributionGroupMember -Identity $addr -Member $upn -BypassSecurityGroupManagerCheck -ErrorAction Stop
                                      Write-Log "Added to distribution list: $addr"
                                      $memberships.Add("Distribution list: $addr")
                                  } catch {
                                      if ($_.Exception.Message -match "already a member") {
                                          Write-Log "Already a member of distribution list: $addr"
                                          $memberships.Add("Distribution list: $addr (already member)")
                                      } else {
                                          Write-Log "FAILED to add to distribution list ${addr}: $($_.Exception.Message)" "WARN"
                                          $problems.Add("Distribution list: $addr")
                                      }
                                  }
                              }
                  
                              foreach ($addr in $allShared) {
                                  try {
                                      Add-MailboxPermission -Identity $addr -User $upn -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false -ErrorAction Stop | Out-Null
                                      Add-RecipientPermission -Identity $addr -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                                      Write-Log "Granted Full Access + Send As on shared mailbox: $addr"
                                      $memberships.Add("Shared mailbox: $addr (Full Access + Send As)")
                                  } catch {
                                      Write-Log "FAILED to grant access to shared mailbox ${addr}: $($_.Exception.Message)" "WARN"
                                      $problems.Add("Shared mailbox: $addr")
                                  }
                              }
                          }
                      }
                  
                      # ==================================================================
                      # FINISH
                      # ==================================================================
                      Write-Log "Onboarding finished for $displayName. Memberships applied: $($memberships.Count). Problems: $($problems.Count)."
                  
                      if ($problems.Count -gt 0) {
                          Send-TeamsAlert -Text "PARTLY DONE: $displayName ($JobTitle, $Department) was created ($upn) but these items need manual follow-up: $($problems -join '; ')" -Color "FFA500"
                      } else {
                          Send-TeamsAlert -Text "SUCCESS: onboarding complete for $displayName ($JobTitle, $Department). UPN: $upn. Start date: $($StartDate.ToString('yyyy-MM-dd'))."
                      }
                  
                      Write-Host ""
                      Write-Host "===================== SUMMARY =====================" -ForegroundColor Cyan
                      Write-Host " Memberships applied:"
                      if ($memberships.Count -eq 0) { Write-Host "   (none)" }
                      foreach ($m in $memberships) { Write-Host "   + $m" -ForegroundColor Green }
                      if ($problems.Count -gt 0) {
                          Write-Host " NEEDS MANUAL FOLLOW-UP:" -ForegroundColor Yellow
                          foreach ($p in $problems) { Write-Host "   ! $p" -ForegroundColor Yellow }
                      }
                      Write-Host " Log file: $logFile"
                      Write-Host ""
                      Write-Host "================== NEW ACCOUNT LOGIN ==================" -ForegroundColor Green
                      Write-Host " Name      : $displayName"
                      Write-Host " Username  : $samAccountName"
                      Write-Host " Email/UPN : $upn"
                      Write-Host " Temp pass : $tempPassword" -ForegroundColor Yellow
                      Write-Host " The user must change it at first logon."
                      Write-Host " It is shown ONCE and is NOT saved in the log. Copy it now."
                      Write-Host "=======================================================" -ForegroundColor Green
                  }
                  catch {
                      Write-Log "ONBOARDING FAILED: $($_.Exception.Message)" "ERROR"
                      if ($adCreated) {
                          Write-Log "NOTE: the AD account '$samAccountName' WAS created before the failure. Check it in AD before running this script again for the same person." "WARN"
                      }
                      Send-TeamsAlert -Text "FAILED: onboarding for $FirstName $LastName. Error: $($_.Exception.Message)" -Color "FF0000"
                      throw
                  }
                  finally {
                      try { Disconnect-MgGraph | Out-Null } catch { }
                      try { if ($needExchange) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } } catch { }
                  }
- edit file to include


| Setting	| Value |
|---|---|
| $UpnSuffix	| contoso.com" |
| $Domain	| "corp.contoso.com" |
| $targetOU	| "OU=$Department,OU=Employee,DC=corp,DC=contoso,DC=com" |
| $LicenseSkuId |	the SkuId from step 8 |
| Entra Connect server | "AAD-01" (once confirmed) |
| Groups section	| Edit to fit org |
| Add-Type -AssemblyName System.Web	added at the top, under the param(...) block
| Update-MgUser ... -UsageLocation "NG"	added just before the Set-MgUserLicense line

<img width="981" height="498" alt="image" src="https://github.com/user-attachments/assets/64b16669-840d-4508-a21b-da928d0d24ea" />

- Check the script for typos
- Open PowerShell and run:

      $errs=$null; [void][System.Management.Automation.Language.Parser]::ParseFile("C:\Scripts\New-EmployeeOnboarding.ps1",[ref]$null,[ref]$errs); $errs
      $errs=$null; [void][System.Management.Automation.Language.Parser]::ParseFile("C:\Scripts\New-EmployeeOnboarding.ps1",[ref]$null,[ref]$errs); $errs | Select-Object @{n='Line';e={$_.Extent.StartLineNumber}}, Message | Format-Table -Wrap
-To remove errors, run:

                  $p = "C:\Scripts\New-EmployeeOnboarding.ps1"
      Copy-Item $p "$p.bak"
      
      $t = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
      $t = $t -replace '[\u201C\u201D]', '"'
      $t = $t -replace '[\u2018\u2019]', "'"
      $t = $t -replace '[\u2013\u2014]', '-'
      $t = $t -replace '[^\x00-\x7F]', ''
      [IO.File]::WriteAllText($p, $t, (New-Object Text.UTF8Encoding($true)))

  <img width="1899" height="711" alt="image" src="https://github.com/user-attachments/assets/4037d603-7c21-46d5-b6fe-04be71c792f3" />


# 8: Make sure the account running it has rights
- In Active Directory Users and Computers, right-click Linkorg-Users, then Delegate Control.
- Add the account you'll run the script with.
- Choose Create, delete, and manage user accounts, then finish.
If you're logged in as a Domain Admin, you can skip this for the test.


# 9: Allow the script to run
powershell

     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

- This only lasts for that window, so it's safe.
<img width="1689" height="257" alt="image" src="https://github.com/user-attachments/assets/16063066-7210-465f-893e-85b74fbc00e4" />



# 10: Run the test with a fake user

- Pick a manager who exists in AD. Use their login name (the sAMAccountName), not their full name. Use a department that matches an OU exactly:

powershell

            Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
            $secureSecret = Import-Clixml -Path "C:\Secrets\onboarding-secret.xml"
             
            C:\Scripts\New-EmployeeOnboarding.ps1 `
                -FirstName "Testing" -LastName "Onboarding" -Department "FINANCE" `
                -Manager "tolowoniyi" -JobTitle "Test Account" -StartDate (Get-Date) `
                -M365Groups "Finance@contoso.com" `
                -DistributionLists "Employees@contoso.com" `
                -SharedMailboxes "it@contoso.com","Example@contoso.com" `
                -ExchangeAdminUpn "admin@contoso.com" `
                -TenantId "uguhjioikojuuhikojggyujiokhuuiguygyg" `
                -ClientId "hbhvghmhjkmjhvhjmvghjkjhhghhggbbhhjh" `
                -ClientSecret $secureSecret

- Replace the three uppercase values. It takes about 5 minutes. Don't close the window. You'll see lines like "AD account created", "Delta sync triggered", "Sync check attempt 1/10", and finally "License assigned" and "Onboarding completed successfully"
<img width="536" height="314" alt="image (3)" src="https://github.com/user-attachments/assets/5e78e85d-aa36-4190-baeb-1cb3b3023790" />

<img width="1029" height="819" alt="image" src="https://github.com/user-attachments/assets/a5a48ba1-e3d1-4639-aec4-7ae18ef7bbf5" />

<img width="481" height="220" alt="image (2)" src="https://github.com/user-attachments/assets/929752f0-53c4-4017-b77b-9cc5543fe47e" />

<img width="302" height="268" alt="image" src="https://github.com/user-attachments/assets/807a2c1f-2e3c-4d5c-82ed-a1d46304cabe" />




# 11: Verify
- Active Directory: tonboarding is in Contoso-Users > FINANCE
- Entra admin center, Users: the user appears with a Business license
- Log file: open the newest file in C:\Logs\Onboarding\ and confirm there are no ERROR lines
<img width="561" height="93" alt="image (4)" src="https://github.com/user-attachments/assets/e4ce9e18-56ff-4ec5-a038-3d60f1a529f2" />

<img width="403" height="396" alt="image (7)" src="https://github.com/user-attachments/assets/9c739816-9fe4-4331-9d4d-a7e01f749fa0" />

<img width="675" height="126" alt="image" src="https://github.com/user-attachments/assets/7be8a16b-cf9c-4e5c-b90d-6f7ba74d0a20" />




# 12: Clean up the test user
- Delete test.onboarding in Active Directory.
- Run Start-ADSyncSyncCycle -PolicyType Delta on the Entra Connect server.
- In the Entra admin center, the user moves to Deleted users, and the license is freed up.
