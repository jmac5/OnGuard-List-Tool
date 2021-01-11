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

# Directions

**Step 1**: Change the server declarations in the beginning of the script to match your organization's settings. If your OnGuard DB server was named "ONGUARD-DB.Company.domain", it would likely be "ONGUARD-DB\LENEL". If you get an error, try the FQDN instead "ONGUARD-DB.COMPANY.DOMAIN\LENEL" 

You will also need the ACCESSLVLID of the level you want to use. First, create an access level with no readers in it, and name it something distinct so it can be found easily. Something like "0" or "1" will go straight to the top, so these are good choices. To find the ACCESSLVLID, look in the database (AccessControl.dbo.ACCESSLVL) and it should be listed. For a name of "0", the query to find it in MS SQL would be "SELECT * FROM AccessControl.dbo.ACCESSLVL WHERE DESCRIPT like '0'". In the example below, this gives us an ACCESSLVID of 40, which we then put in the script in the place of '0000'. 

![Example7](https://github.com/jmac5/OnGuard-List-Tool/blob/main/Tool%20Screenshots/Screen%20Shot%202021-01-11%20at%208.32.35%20AM.png)

![Example2](https://github.com/jmac5/OnGuard-List-Tool/blob/main/Tool%20Screenshots/Screen%20Shot%202021-01-11%20at%208.12.06%20AM.png)

**Step 2**: Once that is all set, you can save the script as whatever you'd like to your machine. This should make it possible to run (running directly after download won't work, and throws a signing error). Once saved, go ahead and "Run with Powershell" or simply double-click the script. If Sql-cmd modules have been installed, it should fire right up. Here's an example of what it looks like when I add a list of 3 IDs, and add our placeholder (0) to them using the program.

**Step 3**: Launch System Administration, and search for anyone who has "0" on their badge. If done correctly, this should yield only the 3 we selected, as shown in the screenshots below. From there, you can do Cardholder > Bulk operations on the list, which should save you significant time! 

![Example4](https://github.com/jmac5/OnGuard-List-Tool/blob/main/Tool%20Screenshots/Screen%20Shot%202021-01-11%20at%208.22.35%20AM.png)

![Example5](https://github.com/jmac5/OnGuard-List-Tool/blob/main/Tool%20Screenshots/Screen%20Shot%202021-01-11%20at%208.23.18%20AM.png)

![Example6](https://github.com/jmac5/OnGuard-List-Tool/blob/main/Tool%20Screenshots/Screen%20Shot%202021-01-11%20at%208.24.16%20AM.png)
