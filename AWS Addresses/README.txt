This program creates an Alias of AWS addresses which can be imported into a WatchGuard firewall.

Process:
downloads the latest version of the AWS ip-ranges json file,
Waits 3 seconds for possible download lag,
Give options to select individual regions with associated service types.

After each iteration of the loop, there is an option to add another region/service pair or to end the loop and generate the alias txt file.
