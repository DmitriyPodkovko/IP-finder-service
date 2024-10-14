Installation:
1. Open the command line (enter 'cmd' in the search).
2. Check whether Python is available (execute 'python --version' in the command line). If not, then install Python.
3. Update the pip package manager (run 'pip install --upgrade pip')
4. In the command line, go to the IP-finder-service directory
5. Execute 'pip install -r requirements.txt'
6. Download Oracle Instant V19 from https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html
7. Unzip to C:\oracle\instantclient_19_24 and check this path in the file IP-finder-service\db\executor.py

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
For offline installation, in our case for Windows Storage Server 2012 R2,
we should install Python 3.10 and run the offline command:
    py -3.10 -m pip install --no-index --find-links ./packages -r requirements.txt

check installed packages:
    py -3.10 -m pip list
   or
    pip list
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
