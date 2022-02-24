import codecs
# import os
wifi_ssid = input("Please enter the name of the Wi-fi being used: ")
wifi_pass = input("Please enter the password: ")

byted = str.encode(wifi_ssid)
hexed = codecs.encode(byted, 'hex')
final = hexed.decode()
xml_output = f'''<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
	<name>{wifi_ssid}</name>
	<SSIDConfig>
		<SSID>
			<hex>{final}</hex>
			<name>{wifi_ssid}</name>
		</SSID>
	</SSIDConfig>
	<connectionType>ESS</connectionType>
	<connectionMode>manual</connectionMode>
	<MSM>
		<security>
			<authEncryption>
				<authentication>WPA2PSK</authentication>
				<encryption>AES</encryption>
				<useOneX>false</useOneX>
			</authEncryption>
			<sharedKey>
				<keyType>passPhrase</keyType>
				<protected>false</protected>
				<keyMaterial>{wifi_pass}</keyMaterial>
			</sharedKey>
		</security>
	</MSM>
	<MacRandomization xmlns="http://www.microsoft.com/networking/WLAN/profile/v3">
		<enableRandomization>false</enableRandomization>
		<randomizationSeed>32143706</randomizationSeed>
	</MacRandomization>
</WLANProfile>'''

xml_write = open("wifi.xml", "w")
xml_write.write(xml_output)
xml_write.close()

Instructions = '1. Open the Advanced Remote Background.\n2. Upload the XML file to the target computer at "C:\\".\n3. run: netsh wlan add profile filename="C:\wifi.xml"\n4. Delete the file.'
help_write = open("Instructions.txt", "w")
help_write.write(Instructions)
help_write.close()