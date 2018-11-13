# Volunter_Scheduler
Automated volunteer schedule for use by Yosemite National Park volunteer coordinator staff.

## Dependencies
* Python 3 
  * Pandas
  * xlwings
* Excel

## Setup
* Install the required dependencies.<br />
  * Install and setup the latest version of [Python3](https://www.python.org/downloads/). <br />
  * Install pip and xlwings for your version of Python3. [(How to install packages)](https://packaging.python.org/tutorials/installing-packages/)
* Download the Volunteer_Scheduler.py and Volunteer_Scheduler.xlsm and place them within the same folder.

## Usage
* Open Volunteer_Scheduler.xlsm
* Select the range where you want to create the schedule
* Click the "Select Schedule Range" button to create the schedule range. The range will persist throughout saved workbooks and will be replaced if the "Select Schedule Range" button again.
* Enter the user inputted parameters in the dark green highlighted cells to the left.
 * Max Work Week: Sets the maximum number of days a volunteer will be allowed to work consecutively.
 * Min Work Week: Sets the minimum number of days a volunteer will be allowed to work consecutivbely.
 * Number of Sites: Sets the number of sites as represented as a number 1-N that a volunteer can be scheduled to.
NOTE: Num Volunteers and Num Days will be autogenerated based on the schedule range selection and should not be changed.
* Press the "Generate New Schedule" button to automatically generate a new schedule. Depending on your computer, the number of volunteers, and the parameters, it may take ~5-10 seconds to process.

## Notes
**How the program prioritizes who works and who doesnt:**
1) Make sure that every site is filled every day.
2) Make sure that overworked volunteers are given a weekend (max number of working days).
3) Make sure that volunteers who have just had a day off will be given another day off to create a 2 day weekend.
4) Make sure that volunteers have a minimum number of days working between weekends (min number of working days)
5) Make sure that all volunteers have the same number of total days off.
6) Prioritize assigning work sites to volunteers who have worked it the fewest number of times.
