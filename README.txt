To run the script, the “Addresses File.xlsx” and the “Formatted Addresses File.xlsx” must be in the same directory as the script.

The cleaning and processing function readAndCleanXlsx() has been commented on line 223. This is because it requires a google cloud api key with geocode and google distance matrix enabled. I have removed mine for safety reasons, but if you have one you can enter it on line 10 and uncomment the function on line 223 to run it.

Currently the script uses the already cleaned “Formatted Addresses File.xlsx”  to connect to a Postgres database and write the entries to it.
