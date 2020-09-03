# calendar-sync

This program is designed to sync a Excel spreadsheet to your Google calendar. It uses colors to determine the appointment type (online, campus, exam or holiday). It is written specifically for Leiden University and will automatically use different times for Online and Campus appointments.

To use this software you need to [create a Google Project](https://console.cloud.google.com/cloud-resource-manager). After this you need to [enable the Calendar API](https://support.google.com/googleapi/answer/6158841?hl=en). Finally [download the credentials](https://cloud.google.com/docs/authentication/getting-started) and put them in same directory as `main.py`, the file should be called `credentials.json`.

In `config.py` you should update the value of `CALENDAR_ID` to match your own Calendar, this can either be `PRIMARY` or another calendar specified by `something@group.calendar.google.com`. To find the url of your calendar go to settings on Google Calendar. You should also specify a relative or absolute path to your Excel spreadsheet (`SCHEDULE`).

To sync your calendar run `python main.py`. This will parse the specified spreadsheet and upload the appointments to your Google calendar. If you make any changes in the spreadsheet just run the script again and it will update your Google calendar.

This project is currently not designed for ease of use. In the future it might become easier to configure and use. You should be aware that there might be bugs and so it is always recommended to let the script sync with a dedicated calendar. That case you can easily remove the calendar if anything gets really messed up without loosing anything else.
