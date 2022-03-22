# Commonly Used Commands
## This file contains some of the most commonly used commands and tools which I use daily in my work.


---

### Powershell

- **_Synchronization service_**
    - `Start-ADSyncSyncCycle -PolicyType Delta`

- **_Domain Join_**
    - `Add-Computer -DomainName domain_name -Credential domain\user`

- **_Repair Domain Connection_**
    - `Test-ComputerSecureChannel -Repair -Credential domain\user`


---

### CMD

- **_General Windows errors_**
    - `sfc /scannow`
    - `dism /online /cleanup-image /restorehealth`

- **_Create new user_**
    - `net user /add User_Name Password`

- **_Add to Admins_**
    - `net localgroup administrators user_name /add`

- **_Enable RDP_**
    - `reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f`
    - `netsh advfirewall firewall set rule group="remote desktop" new enable=Yes`

- **_Add RDP user_**
    - `net localgroup "Remote Desktop Users" User_Name /add`

- **_Connectivity issues_**
    - `ipconfig /flushdns`
    - `ipconfig /release`
    - `ipconfig /renew`

- **_Update Office_**
    - `"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user forceappshutdown=true`

- **_Windows Update_**
    - `wuauclt.exe /detectnow /updatenow`

- **_Retrieve Events_**
    - `wevtutil qe System "/q:*[System[TimeCreated[@SystemTime>='2022-03-03T15:30:00' and @SystemTime<='2022-03-04T22:00:00']]]"`

- **_Checking hash_**
    - `certutil -hashfile file_path encryption_type`

- **_Elevate_**
    - `runas /user:machine_or_domain\user_name cmd`


---

### Printer

 - **_Clear Print Spooler_**
    1. run `services.msc`
    2. stop printer spooler
    3. run `%WINDIR%\system32\spool\printers`
    4. delete contents of the printers directory
    5. start printer spooler

- **_Configure printer driver for RDP printing_**
    1. Launch ‘gpedit.msc’ from the ‘Run’ command
    2. Open Computer Config -> Admin Templates -> Windows Components -> Remote Desktop Services -> RD Session Host -> Printer Redirection
    3. Edit the setting: Use Remote Desktop Easy Print printer driver first
    4. Change this to ‘Disabled’

- **_Another resource for RDP printing_**
    - https://www.farmhousenetworking.com/networking/remote-access/remote-desktop-network-printer-redirection/


---

### Registry

- **_If the HD is detected as a USB drive"_**
    - Navigate in regedit to `HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control`
    - Change PortableOperatingSystem to 0

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

### Web

- **_Watchguard download:_**
    - https://cdn.watchguard.com/SoftwareCenter/Files/MUVPN_SSL/12_7_2/WG-MVPN-SSL_12_7_2.exe

- **_Adobe Downloads_**
    - https://get.adobe.com/reader/



---

### Other

- **_Scan PST_**
    - Navigate to `C:\Program Files\Microsoft Office\root\Office16`
    - run scanpst.exe

- **_Fix Blurry Screen_**
    - Search for "Adjust ClearType Text"

- **_To get a computer out of S Mode_**
    1. Settings->Update&Security->Activation->Go to the Store
    2. Select "Get" under "Switch out of S Mode"



---

### Remote access on Apple devices
   - **_For iOS_**
       1. Download OpenVPN Connect
       2. Download the .ovpn file
       3. Goto Downloads and open the .ovpn file in OpenVPN Connect
       4. Enter the required fields
       5. Turn on VPN
       6. Files>...>Connect to Server>enter local address of the server

   - **_For MacOS_**
        - https://setapp.com/how-to/map-a-network-drive-on-mac