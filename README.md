# calendar-sync

This program is designed to sync a Excel spreadsheet to your Google calendar. It uses colors to determine the appointment type (online, campus, exam or holiday). It is written specifically for Leiden University and will automatically use different times for Online and Campus appointments.

To use this software you need to [create a Google Project](https://console.cloud.google.com/cloud-resource-manager). After this you need to [enable the Calendar API](https://support.google.com/googleapi/answer/6158841?hl=en). Finally [download the credentials](https://cloud.google.com/docs/authentication/getting-started) and put them in same directory as `main.py`, the file should be called `credentials.json`.
