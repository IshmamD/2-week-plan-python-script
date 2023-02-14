# 2-week-plan-python-script
Explore Maximo work orders from 1 date to another in the future and gather their info
This is an exact copy of the documentation I placed for the person who would go on to use this script

---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

1. Before use, make sure to have some version of python 3, and the python packages selenium and xlwings. These are for interacting with the browser and with excel respectively. 
Python comes with a built in IDE called IDLE. Use IDLE instead of typing python into the search bar and launching that, because thats basically command prompt but with python.
You must use IDLE to launch these python scripts. I tried another IDE (Pycharm) and it didnt work.

2. To show line numbers in IDLE, go to options > Show line numbers
2.1 - Before doing anything more, go to the username and password fields in the code and enter yours. These are on line 22 and 24

3. The query is completed automatically in this script, but you can change the query if you wish by viewing lines 53, 57, 60, and 65 to change the query for already existing
text boxes that are filled by the script. read the variable name to see which box is which. If you want to select an entirely new text box for your query then
more detailed info can be found at https://automatetheboringstuff.com/2e/chapter12/
Scroll down to the part "Controlling the Browser with the selenium Module" and read from there if needed.

4. Make sure the script and the excel file are in the same directory/folder

5. If your computer is slow or busy, the script may crash. If it does, it will tell you the line on which it crashed. If it does that, before that line type in time.sleep(2), 
and it should fix it. If issues continue, increase the number from 2 to a higher number to increase the wait time.
2 represents 2 seconds in this case. Try to use the lowest wait time possible to increase time efficiency since this script is already pretty slow.

6. To change the time period that the script looks through, look at line 34 and change the time delta from "(days = 14)" to any other number. it can also be measured
in months, for example "(months = 2)", weeks, years, or even seconds.

7. The python script does not do the total hours since this is trivial to do in excel. Do not touch the Excel spreadsheet except for the scrollbar while the script is running
It will crash.
7.1. In the excel sheet, any hour section labeled "0" should be changed to 0.5.

8. The script may appear slow. That's because it is. The script can only collect data as fast as the webpage will load, but sometimes it loads slowly so to ensure the
script doesn't crash, it has wait times set to be reasonably slow. You shouldn't close/restart the script unless IDLE gives you an error, or something doesn't happen either
in the browser or in Excel for 30 seconds or more.

9. If there's an error right before the query making step that says "stale element reference: element not attached to page", just run it again.

10. by my calculations the script takes 1 hour and 10 minutes to complete its tasks. Leave it running in the background, and make sure your PC doesn't go to sleep while its
running the script. It's best to use this script while you're doing desk work, or if you're only going to be away from your PC for short periods. The monitor turning off
and log outs do not affect the script, only sleep.
