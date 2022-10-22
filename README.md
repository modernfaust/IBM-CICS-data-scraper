# IBM-CICS-data-scraper

A custom VBA script that scrapes data from a proprietary IBM CICS database. 

This was created for my former employer to automate a highly repetitive and error prone function.

The Objective:

Formerly, a production order report required an hour of preparation. Data had to be manually extracted and verified from the CICS database via an Attachmate Reflections terminal emulator. 
Thankfully, Attachmate Reflections provided an API that could automate data collection.

My Work:

1. Classify Data
- After collecting the data, I wrote a subroutine that could identify and usefully classify the data according to business needs. Different types of production orders had to be prepared according to certain situations. This functionality was written in scrapeWorkOrder.

2. Format Work order
- Production orders needed to be stylistically consistent. Thus I created a button that could guarantee consistency throughout the pages.

3. Compliance
- Prior to completing the work order, the preparer must stamp and save the file as both PDF and XLSM according to naming conventions. I created a button to do this.

4. Revise and Iterate
- Several data fields were missing or incorrectly picked up. Thus I made hundreds of lines of revisions to guarantee that data was being pulled from the correct locations.

5. Analysis
- Categorize and generate statistics for other departments

