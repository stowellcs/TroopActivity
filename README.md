# Troop Activity Parser
## Overview
The Troop is using Square to collect payments for Scouts and Scouters for campouts. Transaction data in a csv file is downloaded periodically to see who has signed up. 

In the transaction file, there is a "Modifiers Applied" field that contains all responses to questions, but the answers to the questions are often found in different positions in this field varies and it can take a lot of manual effort to parse it for presentation to the campout coordinators.
 
This Python code can be used to quickly parse that modifiers field.  It takes the csv file downloaded from the Squareup.com web site and provides an Excel-formatted output.

## Dependencies
To begin using this program, make sure you have Python 3.9 and have installed a couple packages, including pandas1.5.2 and Xlsxwriter 3.0.7.


## Process
Once you are set up, download the Transactions.csv from this site: https://squareup.com/login?return_to=%2Fdashboard%2Fsales%2Ftransactions

Before running the Python code, be sure to change the directory paths and file names in the code.
  *myDirectory - where the csv file is located
  *fileName - name of csv file
  *filterChoice1 - **Very important!**  Use this to match only those downloaded transactions for the event you want.  For example, the csv file may have transactions with "Item" descriptions as "Scout Registration - Ropes Course Campout - May 2023" or "Scouter Registration - Ropes Course Campout - May 2023".  Enter the phrase after the Scout/Scouter Registration.  Example: "Ropes Course Campout - May 2023"
  *filterChoice2 - optionally, to report on two events at once, fill in this variable with the same rules as filterChoice1.
  *eventName - this phrase is used to label the output files.  For example, if the variable is assigned a value = 'May 2023 Ropes Course', the file output will start with that same phrase.
  
  Contact: christopher_stowell@yahoo.com if you have questions.
