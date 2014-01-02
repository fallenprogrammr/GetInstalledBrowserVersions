GetInstalledBrowserVersions
===========================

Powershell script to get the installed versions of web browsers on local and remote machines.

The script looks for a "crossbrowser.xls" excel sheet [This will probably change in the future] and enumerates through the first column starting with row #2, treating the values as names of machines the script has to check the browser versions on.

The script then attempts a ping to that machine and if the machine is reachable, it attempts to get the browser versions of IE, Chrome and Firefox [Opera support will be added later] installed on the machine.

The output is piped to a list.csv file in the same directory.

If no "crossbrowser.xls" is found, then the local machine is used for checking the browser versions.

Future todos:

Make input and output files parameterized.

Replace excel source type with an alternative.

Add opera support for checking browser version.

Add information if a browser cannot be detected as an output to [list.csv].
