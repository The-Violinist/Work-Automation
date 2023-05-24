# Commonly Used Commands
## This file contains some of the most commonly used commands and tools which I use daily in my work.


---
### CMD

- **_General Windows errors_**
    - `sfc /scannow`
    - `dism /online /cleanup-image /restorehealth`

- **_Create new user_**
    - `net user /add User_Name Password`

- **_Add to Admins_**
    - `net localgroup administrators user_name /add`

- **_Get Public IP_**
    - `nslookup myip.opendns.com. resolver1.opendns.com` or
    - `curl http://ipecho.net/plain`

- **_Retrieve Wifi Password_**
    - If unsure of the SSID: `netsh wlan show profiles` then use that SSID in the following command:
    - `netsh wlan show profile name="SSID_Name" key=clear`
- **_Enable RDP_**
    - `reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f`
    - `netsh advfirewall firewall set rule group="remote desktop" new enable=Yes`

- **_Add RDP user_**
    - `net localgroup "Remote Desktop Users" User_Name /add`

- **_Restart a service_**
    - In this instance, I am using the N-Able Take Control service
    - `net stop "BASupportExpressStandaloneService_LOGICnow" && net start "BASupportExpressStandaloneService_LOGICnow"`

- **_Find and Shutdown Process_**
    - `tasklist | findstr /I "KEYWORD"`
    - `taskkill /f /pid pid_num_from_last_command`

- **_Connectivity issues_**
    - `ipconfig /flushdns`
    - `ipconfig /release && ipconfig /renew`

- **_Elevate_**
    - `runas /user:machine_or_domain\user_name cmd`

- **_Update Office_**
    - `"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user forceappshutdown=true`

- **_Windows Update_**
    - `wuauclt.exe /detectnow /updatenow`

- **_Retrieve Events_**
    - `wevtutil qe System "/q:*[System[TimeCreated[@SystemTime>='2022-03-03T15:30:00' and @SystemTime<='2022-03-04T22:00:00']]]"`

- **_Checking hash_**
    - `certutil -hashfile file_path encryption_type`

- **_Get Printer List_**
    - `wmic printer list brief`

- **_Logoff a user_**
    - Get session ID `query session`
    - Logoff the user `logoff <session ID>`

- **_Find Service belonging to port_**
    - Get process ID: `netstat -aon | findstr "<port number>"`
    - Use the process ID found: `tasklist /fi "PID eq <ID>`
---

### Powershell

- **_Synchronization service_**
    - `Start-ADSyncSyncCycle -PolicyType Delta`

- **_Domain Join_**
    - `Add-Computer -DomainName domain_name -Credential domain\user`

- **_Repair Domain Connection_**
    - `Test-ComputerSecureChannel -Repair -Credential domain\user`

- **_Get Public IP_**
    - `Invoke-WebRequest ifconfig.me/ip`

- **_Get AD Password Requirements_**
    - `Get-ADDefaultDomainPasswordPolicy`
---

### Printer

 - **_Clear Print Spooler_**
    - run `services.msc`
    - stop printer spooler
    - run `%WINDIR%\system32\spool\printers`
    - delete contents of the printers directory
    - start printer spooler

- **_Configure printer driver for RDP printing_**
    - Launch ‘gpedit.msc’ from the ‘Run’ command
    - Open Computer Config -> Admin Templates -> Windows Components -> Remote Desktop Services -> RD Session Host -> Printer Redirection
    - Edit the setting: Use Remote Desktop Easy Print printer driver first
    - Change this to ‘Disabled’

- **_Another resource for RDP printing_**
    - https://www.farmhousenetworking.com/networking/remote-access/remote-desktop-network-printer-redirection/


---

### Registry/Group Policy

- **_If the HD is detected as a USB drive_**
    - Navigate in regedit to `HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control`
    - If PortableOperatingSystem is present in the Control folder:
        - Change PortableOperatingSystem value to 0

- **_Turn on/off USB access_**
    - Navigate in gpedit to `Computer Configuration\Administrative Templates\System\Removable Storage Access`

- **_Disable Cortana_**
    - Navigate in regedit to `HKEY_Local_Machine\SOFTWARE\Policies\Microsoft\Windows`
    - Create a new key called `Windows Search`
    - Create a new DWORD (32 bit) called `AllowCortana` and set its value to 0

- **_Disable Bing_**
    - Navigate in regedit to `HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows`
    - Create a new key called `Explorer`
    - Create a new DWORD (32 bit) called `DisableSearchBoxSuggestions` and set its value to 1

---

### Other

- **_Scan PST_**
    `C:\Program Files\Microsoft Office\root\Office16`

- **_Fix Blurry Screen_**
    - Search for "Adjust ClearType Text"

- **_To get a computer out of S Mode_**
    - Settings->Update&Security->Activation->Go to the Store
    - Select "Get" under "Switch out of S Mode"



---

### Remote access on Apple devices
- **_For iOS_**
    - Download OpenVPN Connect
    - Download the .ovpn file
    - Goto Downloads and open the .ovpn file in OpenVPN Connect
    - Enter the required fields
    - Turn on VPN
    - Files>...>Connect to Server>enter local address of the server

- **_For MacOS_**
    - https://setapp.com/how-to/map-a-network-drive-on-mac
