# ADUsers

Command-line tools for fetching users from Active Directory and save it to CSV file. Just for C# example.

# Usage

**adusers.exe [options]**
Where options as follow:
* -help  This screen.
* -CN:   Common Name. Optional parameter and may be skipped.
* -OU:   Organisation Units, separated by a point.
* -DC:   Full Domain Controller Name, for example: DC:dc1.example.com
* -out:<out.csv> -- Optional: specify output file name. 'ADUsers.csv' by default.
* -usr:<mask> -- An username mask to search for. Default: '*'. Example: '-usr:50*'

# AUTHOR
   An0ther0ne

