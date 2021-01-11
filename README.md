# OnGuard-List-Tool

A tool written to interact with a SQL Database to upload a list of users as a CSV to apply access to via the OnGuard Access Control System. The actual application has no mechanism for applying access to an otherwise unrelated subset of the population, and this program/script allows for the selection of a list of users via their ID numbers, automatically sorts them into a CSV, and applies an empty access level which can be searched for inside the application. This script/program has reduced the time needed to apply access to cardholders by an order of magnitude, at least.  

PRE-REQ: SqlServer cmdlets for Powershell, available here: https://docs.microsoft.com/en-us/sql/powershell/download-sql-server-ps-module?view=sql-server-ver15

To use the application, simply launch it with PowerShell, and a GUI should show up. From here, you can paste in a list of ID numbers or badge numbers (your choice), and separate them by line. It's ok if Names/emails/ etc get mixed in, they will be stripped out when processed. Once you have modified the script with your servers/access level ID, pressing the ADD PLACEHOLDER TO LIST button will assign that access level to all the badges/IDs on that list. This will allow you to do a search in OnGuard for "Has at least one of the selected levels", and selecting the empty level used as a placeholder. 

![Example](https://github.com/jmac5/Portfolio/blob/master/OnGuard%20DB%20Tool%20Screenshot.png)

Here you can see the following:

* The padded 0s are removed from badge "1"
* Badge 22943 is added without issue
* The random ASCII is ignored, and an error is thrown for invalid IDs
* "hello" is ignored
* "I am not an ID number" is ignored
* The date of 1/8/21 is ignored
* 50975 is added without issue
* Blank line is ignored
* 175940103 is added without issue

The "CSV Verification" shows the final result sent to the DB server, so you can make sure what you sent is what you wanted. In this case, it verifies that only the valid IDs from the input above have been provided to the database query. This was mainly for my debugging purposes, but it can prove helpful or comforting to you, as well. 
