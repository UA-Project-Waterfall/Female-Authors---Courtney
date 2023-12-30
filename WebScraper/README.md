# Lens Scraper (Selenium and PyAutoGUI)
 
This program was made to web-scrape Lens.org for specific inventors

Here are instructions on how to use it.

Add an excel file containing the list of authors to the script folder.
Make sure that there is a header row that has a "First Name" and "Last Name" column.
Run the LensScraper.py file either through the terminal or with a double click.
Follow the prompts to enter the name of the author list excel file.
Alternatively, you can drag the desired excel file into the terminal window to paste its path.
Then enter the speed (seconds per query) at which lens.org will be accessed.
A smaller time (faster querying) may alert the website to detect unusual traffic.

When the program starts, make sure not to move the mouse or press keys until the new windows
that were opened by the program have fully loaded lens.org and bypassed bot detection.
The program will prompt when it is okay to move the mouse again.
