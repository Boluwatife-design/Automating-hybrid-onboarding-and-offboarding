

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
