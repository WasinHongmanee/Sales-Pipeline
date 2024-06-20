# Sales Pipeline

Query sales data using MS SQL from the last 7/30/90 days from my new restaurant’s POS database server and upload to Google Sheets using their API in Python once a week. A cron job is setup to run the python script at 11pm every Sunday night to ingest the data. Automatically refresh the Tableau dashboard every time the Google Sheets updates.
