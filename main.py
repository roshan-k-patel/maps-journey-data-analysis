import googlemaps
import numpy as np
import pandas as pd
from openpyxl.workbook import Workbook
import psycopg2
import xlrd


# Google cloud API key with google distance matrix and geocode enabled
api_key = ""

# Input your database connection info below
DB_NAME = "ilqwhrky"
DB_USER = "ilqwhrky"
DB_PASS = ""
DB_HOST = "john.db.elephantsql.com"
DB_PORT = "5432"

# Takes in a "cleaned" address as input and returns its coordinates and the address in standardised google maps format
def searchAddress(address):
    geocode_result = gmaps.geocode(address)

    lat = geocode_result[0]["geometry"]["location"]["lat"]
    lon = geocode_result[0]["geometry"]["location"]["lng"]
    formatted_address = gmaps.reverse_geocode((lat, lon))[0]["formatted_address"]

    lat_lon = str(lat)[0:8] + ", " + str(lon)[0:8]
    return (lat_lon, formatted_address)

# Uses the google distance matrix library to calculate the distance between two addresses in km
def searchDistance(source, destination):
    distance_query = gmaps.distance_matrix(source, destination)

    distance_parts = distance_query["rows"][0]["elements"][0]["distance"]["text"].split(" ")

    return (distance_parts[0])

# Reads the Addresses File.xlsx provided and performs data cleaning and processing
def readAndCleanXlsx():

    # Pandas library used to write formatted data to new excel sheet, column named listed below
    w = pd.ExcelWriter("Formatted Addresses File.xlsx")
    df = pd.DataFrame(
        columns=["Uid", "Source", "Destination", "Distance(km)", "Source lat/long", "Destination lat/long",
                 "Source State/Territory", "Destination State/Territory"])
    df.to_excel(w, startrow=0, index=False)
    w.save()

    # variables to keep track of row number in unformatted excel file and also in formatted one,
    # they differ due to poor results being excluded from the final formatted dataset
    row = 0
    pd_row = 1

    # while loop to perform the iteration for every row in the sheet, starting from the first entry at row 0
    while row < len(data):

        # empty string placeholders for the data points, these are used for checks so that if the value is still
        # empty after attempted extraction, secondary extraction measures will be attempted
        source = ""
        destination = ""
        distance_km = ""
        src_lat_lon = ""
        dtn_lat_lon = ""
        src_state = ""
        dtn_state = ""

        # blanket try except block to move on to the next data point if an exception is thrown, usually due to bad
        # formatting / out of bounds errors
        try:
            if pd.isnull(data['source_address'][row]):
                source = data['merged_source_address'][row][1:-1].strip()
                destination = data['merged_destination_address'][row][1:-1].strip()
            else:
                source = data['source_address'][row]
                destination = data['destination_address'][row]

            if pd.isnull(data['distance'][row]) or data['distance_units'][row] == "IRRELEVANT":
                distance_km = -1
            else:
                distance_km = str(data['distance'][row])

            # blacklisted terms that are checked against as one of the measures
            # to determine if an address is valid or not
            blacklist = ["IRRELEVANT", "?", "^", "=", "}", "{"]

            # boolean value to keep track of if the address is valid and ultimately decide whether to call the api
            valid_address = True

            # checks both addresses for blacklisted terms
            for term in blacklist:
                if source.__contains__(term) or destination.__contains__(term):
                    valid_address = False

            # calls relevant api's and saves / formats the results
            if valid_address:
                source_geocode = searchAddress(source)
                src_lat_lon = source_geocode[0]
                source = source_geocode[1]

                src_address_parts = source.split(",")
                src_state = src_address_parts[len(src_address_parts) - 2][1:-7]

                destination_geocode = searchAddress(destination)
                dtn_lat_lon = destination_geocode[0]
                destination = destination_geocode[1]

                dtn_address_parts = destination.split(",")
                dtn_state = dtn_address_parts[len(dtn_address_parts) - 2][1:-7]

                if src_state == "":
                    src_state = src_address_parts[len(src_address_parts) - 3].strip()

                if dtn_state == "":
                    dtn_state = dtn_address_parts[len(dtn_address_parts) - 3].strip()

            if distance_km == -1 and valid_address:
                distance_km = searchDistance(source, destination)

            print (row)
            #print(source, destination, distance_km, src_lat_lon, dtn_lat_lon, src_state, dtn_state, sep='\n')
            #print()

            # creates a dataframe and writes the collected info to file while also incrementing the row pointer
            if valid_address:
                # Create dataframe and write to file
                df = pd.DataFrame(
                    columns=[pd_row, source, destination, distance_km, src_lat_lon, dtn_lat_lon, src_state, dtn_state])
                df.to_excel(w, startrow=pd_row, index=False)
                pd_row += 1
                w.save()

        except:
            pass

        # increments the row pointer on the unformatted excel file
        row += 1

    w.close()

# function to connect to a postgres database using the global variables at the beginning of the script
def connectdb():
    try:
        conn = psycopg2.connect(database=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST, port=DB_PORT)

        print("Database connected successfully")
        return conn
    except Exception as e:
        print("Database not connected")
        print(e)

# function that creates a table A for the storage of the data
def createTable(conn):
    try:
        cur = conn.cursor()
        cur.execute("""

        CREATE TABLE A
        (
        UID INT PRIMARY KEY NOT NULL,
        SOURCE TEXT NOT NULL, 
        DESTINATION TEXT NOT NULL,
        DISTANCE_KM TEXT NOT NULL,
        SOURCE_LAT_LON TEXT NOT NULL,
        DEST_LAT_LON TEXT NOT NULL,
        SOURCE_STATE_TERR TEXT NOT NULL,
        DEST_STATE_TERR TEXT NOT NULL
        )

        """)

        conn.commit()
        print("Table Created Successfully")

    except Exception as e:
        print("Error thrown when creating table, possible duplicate")
        print(e)

# function that reads the formatted excel file and writes each row of data to the database
def insertData(conn):
    db_data = pd.read_excel("Formatted Addresses File.xlsx")
    row = 0

    while row < len(db_data):

        uid = db_data['Uid'][row]
        source = db_data['Source'][row]
        destination = db_data['Destination'][row]
        distance_km = db_data['Distance(km)'][row]
        src_lat_lon = db_data['Source lat/long'][row]
        dtn_lat_lon = db_data['Destination lat/long'][row]
        src_state = db_data['Source State/Territory'][row]
        dtn_state = db_data['Destination State/Territory'][row]

        # Try catch block incase the relation already exists or some other error
        try:
            cur = conn.cursor()

            # Inserts data
            cur.execute(
                'INSERT INTO A (UID, SOURCE, DESTINATION, DISTANCE_KM, SOURCE_LAT_LON, DEST_LAT_LON, SOURCE_STATE_TERR,DEST_STATE_TERR) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)',
                (int(uid), source, destination, distance_km, src_lat_lon, dtn_lat_lon, src_state, dtn_state))

            conn.commit()

        except Exception as e:
            # Rolls back a failed query after trying to commit whatever may have been left in the buffer
            print(e)
            conn.commit()
            cur.execute("rollback")

        row += 1

    conn.close()


if __name__ == '__main__':
    data = pd.read_excel("Addresses File.xlsx")

    # creates the api connection using the api key in the global variables at the beginning of the script
    gmaps = googlemaps.Client(key=api_key)

    # Reads and cleans the original data given. This can be uncommented and run if you enter a value for the api key on line 10
    #readAndCleanXlsx()

    # These remaining functions require the Formatted Addresses File.xlsx produced from the readAndCleanXlsx() function
    # to be in the same directory as the script

    # Connects to db using global variables at the start of script
    conn = connectdb()

    # Creates a table A to store values
    createTable(conn)

    # Reads the formatted xlsx file and inserts the data into the postgres database
    insertData(conn)
