# Packet Filter Creator

## Create an importable packet filter on a Watchguard Firewall

This program takes input from the 'port_list' file and creates an importable XML file to create a custom packet filter on Watchguard Firewalls.

The 'port_list' file contains 5 sections as a layered list:
- Enter the new policy name in quotation marks within the outer square brackets (list item 0)
- TCP single ports
    - Entered in the first set of square brackets (list item 1) as comma separated values.
- UDP single ports
    - Entered in the second set of square brackets (list item 2) as comma separated values.
- TCP port ranges
    - Entered in the third set of square brackets (list item 3), within an additional set of square brackets for each range. Beginning and ending ports are entered as comma separated values.
- UDP port ranges
    - Entered in the last set of square brackets (list item 4), within an additional set of square brackets for each range. Beginning and ending ports are entered as comma separated values.

Example: ["Policy_Name",[80],[443],[[1024,65535]],[[1024,65535]]]