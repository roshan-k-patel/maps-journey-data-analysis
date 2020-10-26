To run the script, the “Addresses File.xlsx” and the “Formatted Addresses File.xlsx” must be in the same directory as the script.

The cleaning and processing function readAndCleanXlsx() has been commented on line 223. This is because it requires a google cloud api key with geocode and google distance matrix enabled. I have removed mine for safety reasons, but if you have one you can enter it on line 10 and uncomment the function on line 223 to run it.

Currently the script uses the already cleaned “Formatted Addresses File.xlsx”  to connect to a Postgres database and write the entries to it. You can add your database info from line 12 so that it connects and inserts into your database.

Thank you.

Here are the SQL queries requested:

1.)
SELECT SOURCE,DESTINATION,SOURCE_LAT_LON, DEST_LAT_LON FROM A WHERE SOURCE != NULL and DESTINATION !=NULL;

2.)
select distinct SOURCE_STATE_TERR,SOURCE from A;

3.)
select distinct DEST_STATE_TERR,DESTINATION from A;

4.)
select SOURCE,DESTINATION,DISTANCE_KM from A where SOURCE!=null and DESTINATION!=null;

5a.)
select (total*100)/(select count(*) from A) as percentage,SOURCE_STATE_TERR from (select count(SOURCE_STATE_TERR)as total, SOURCE_STATE_TERR from A group by SOURCE_STATE_TERR) temp;

5b.)
select (total*100)/(select count(*) from A) as percentage,DEST_STATE_TERR from (select count(DEST_STATE_TERR)as total, DEST_STATE_TERR from A group by DEST_STATE_TERR) temp;