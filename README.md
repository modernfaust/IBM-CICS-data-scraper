# IBM-CICS-data-scraper

A custom VBA script that scrapes data from a proprietary IBM CICS database. 

This was created for my former employer to automate a highly repetitive and error prone function.

The Objective:

Formerly, a production order report required an hour of preparation. Data had to be manually extracted and verified from the CICS database via an Attachmate Reflections terminal emulator. 
Thankfully, Attachmate Reflections provided a power scripting API that could automate data collection.
A member of our team wrote a rudimentary yet functional edition of the script that could populate 60% of the fields necessary for the production order.
When I joined, I sought to populate 100% of the fields.

My Work:

1. Deducing the type of Production Order to prepare
- After collecting the data, I wrote a subroutine that could identify and usefully classify the data according to business needs. Different types of production orders had to be prepared according to certain situations. This functionality was written in scrapeWorkOrder.

2. Formatting the page
- Production orders needed to be stylistically consistent. Thus I created a button that could guarantee consistency throughout the pages.

3. Stamp and save
- Prior to completing the work order, the preparer must stamp and save the file as both PDF and XLSM according to naming conventions. I created a button to do this.

4. Missing fields
- Several data fields were missing or incorrectly picked up. Thus I made hundreds of lines of revisions to guarantee that data was being pulled from the correct locations.


