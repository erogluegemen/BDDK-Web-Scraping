# BDDK-Web-Scraping (currently updating..)
- Collecting BDDK(Banking Regulation and Supervision Agency) data from the official website(https://www.bddk.org.tr/).
- The code scheduled with Windows Task Scheduler to run every 7 days.
- *.py* file for pulling the Excel file from the website. (Automation)
- *.ipynb* notebook for pulling the whole table from the website. (Scraping)

## How you can set up a cron job to run your code every 7 days on a Unix-like system:
1. Open a terminal.

2. Type crontab -e and press Enter. This will open the crontab configuration in an editor.

3. Add the following line to the crontab file to schedule your script to run every 7 days:

   ```bash
   0 0 */7 * * /path/to/python /path/to/your/script.py
   ```

  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Replace /path/to/python with the path to your Python interpreter and /path/to/your/script.py with the path to your Python script.

4. Save and exit the crontab editor.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The above cron schedule (0 0 */7 * *) means that the script will run at 12:00 AM (midnight) every 7 days.


## For Windows using Task Scheduler:
1. Open the Task Scheduler application.

2. Click on "Create Basic Task" or "Create Task," depending on the version of Task Scheduler you are using.

3. Follow the prompts to set a name and description for the task.

4. Choose "Daily" trigger and set the "Recur every" option to 7 days.

5. Select "Start a Program" as the action.

6. Browse and select the Python executable as the program/script, and provide the path to your Python script in the "Add arguments" field.

7. Complete the wizard and save the task.

This will create a scheduled task that runs your script every 7 days.
