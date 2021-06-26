# Taxi-Script
Purpose: To reduce the man-time needed to sort the taxi booking data and keying in values.

Company: [Stream Mobility](http://streammobility.com.sg/)

Role: No direct ties with the organisation, just helping a friend in my free time.

Approach: 
- Used Python to create a script for sorting the data using mostly list data structure.
- Utilized openpyxl, xlrd and xlsxwriter modules to format the data in a readable xlsx file for submission.
- Took advantage of the power of iterations to automate simple and mundane sorting tasks.
- Integrated tkinter module to enhance user interface and made the script more user-friendly.

Changelog v1.1
- Added total trips calculation for each day into every sheet when running option 2.
- Added new headers for summary sheet, including entire sheet total trips, prices and days selected.
- Optimized the code such that the user does not need to input sheet name when loading files.
- Fixed a bug causing the prices to lose its number format.
- Added an error message disallowing option 2 to run if there are sheets of the incorrect format.
- New sheets created are now renamed to the exact date instead of '1','2','3'.

Changelog v1.2
- Added a new feature to format online calendar to required format

Changelog v1.3
- Added a file dialog using tkinter
- Removed unnecessary user prompt
- Added combo option (4 followed by 1)
- Added new driver Column J as "Ong"
- Auto-detect file type and disable options accordingly
- Add stream mobility picture
