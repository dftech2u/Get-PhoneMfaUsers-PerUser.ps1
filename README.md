# Get-PhoneMfaUsers-PerUser.ps1
This script lists every Microsoft 365 user in a client tenant who has SMS or voice registered for MFA, and exports the results to CSV. Use it before disabling SMS/voice in the Authentication methods policy, or when auditing a client's MFA posture.

Requirement
Details - Workstation - PC with Windows PowerShell 5.1 (built in) or PowerShell 7

PowerShell module
Microsoft Graph PowerShell SDK (Microsoft.Graph). Install steps below

Admin role in the client tenant
Global Reader or Authentication Administrator, at minimum

Graph permissions
User.Read.All, UserAuthenticationMethod.Read.All, Organization.Read.All, AuditLog.Read.All. The script requests these at sign-in

Consent
First run in a tenant prompts for consent. A Global Administrator may need to approve it if user consent is restricted

Internet access
The workstation must reach login.microsoftonline.com, graph.microsoft.com, and the PowerShell Gallery

Step 1: Install the Microsoft Graph PowerShell module
This is a one-time step per workstation. The script does not check for the module and will not install it for you.
1. Open PowerShell. No admin rights are needed, because the install is for your user account only.
2. Check whether the module is already installed:
   Get-InstalledModule Microsoft.Graph -ErrorAction SilentlyContinue
   If a version number is returned, skip to step 5. If nothing is returned, continue.
3. On Windows PowerShell 5.1, make sure TLS 1.2 is enabled so the PowerShell Gallery can be reached:
   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
4. Install the module:
   Install-Module Microsoft.Graph -Scope CurrentUser
   If asked to install the NuGet provider, answer Y. If warned about an untrusted repository (PSGallery), answer Y or A. The install downloads many sub-modules and can take several minutes.
5. Confirm the modules the script uses are present:
   Get-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -ListAvailable
   You should see all three listed.
To update the module later, run Update-Module Microsoft.Graph.

Step 2: Getting the Script
1. Download Get-PhoneMfaUsers-PerUser.ps1.
2. Move it to a working folder, for example C:\Scripts\MFA-Audit. The CSV reports are saved to whatever folder you run the script from.
3. Unblock the file. Windows marks downloaded files as coming from the internet, and PowerShell may refuse to run them:
   Unblock-File C:\Scripts\MFA-Audit\Get-PhoneMfaUsers-PerUser.ps1
4. If scripts are blocked on the workstation, allow them for the current PowerShell window only:
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   This resets when the window closes and does not change the machine's policy.

Step 3: Running the Script
1. Change to the working folder:
   cd C:\Scripts\MFA-Audit
2. Run the script, pointing it at the client's tenant:
   .\Get-PhoneMfaUsers-PerUser.ps1 -TenantId clientdomain.onmicrosoft.com
3. A Microsoft sign-in window opens. Sign in with an admin account in the client tenant that has Global Reader or Authentication Administrator.
4. On the first run in a tenant, a permissions consent screen appears. Review the four permissions and accept. If you see "Need admin approval," a Global Administrator must consent first.
5. Wait for it to finish. A progress bar shows the user it is on. Expect a few users per second, so a 300-user tenant takes a few minutes.

Optional parameters
Parameter
What it does
-TenantId
Tenant ID or domain to connect to. Always use this when working in a client tenant
-IncludeDisabled
Include disabled accounts (skipped by default)
-IncludeGuests
Include guest/B2B accounts (skipped by default)
-OutputPath
Custom path for the CSV, e.g. -OutputPath C:\Reports\Contoso_MFA.csv
When finished, disconnect so the next run doesn't reuse this tenant's session:
Disconnect-MgGraph
   
