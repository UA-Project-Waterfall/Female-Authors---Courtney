# Lens Scraper (Selenium and PyAutoGUI)
 
This program was made to web-scrape Lens.org for specific inventors

Here are instructions how to use it.

Add an excel file containing the list of authors to the script folder.
Make sure that there is a header row that has a "First Name" and "Last Name" column.
Run the LensScraper.py file either through terminal or the bat file.
When it prompts you, enter the name of the author list file.
The seconds per query is the speed that lens.org is accessed.
A smaller time (faster querrying) may alert the website to detect unusual traffic.

When the program starts, make sure not to move the mouse until the new windows
that were opened by the program have fully loaded lens.org and bypassed bot detection.