# Script to prevent login by anyone who is not me
# More or less a joke script which stops login to a computer if the username is not me.

### LIBRARIES ###
import os
import re

### VARIABLES ###
user = os.getlogin()                                            # Get username
query_output = os.popen(f'query session {user}').read()         # Retrieve session query for the user

### FUNCTIONS ###
def logout():
    ID = int(re.search(r'\d+', query_output).group())           # Get the session ID
    os.system(f"logoff {ID}")                                   # Log off the user with that ID

### MAIN ###
if user == "<Your Unername>":                                   # If the user is you, end the program
    quit()
else:                                                           # Otherwise, log the user out
    logout()

### END ###
