# ManHourTracker_Beta_04-11-2016

Date Written: 04/11/2016

Industry: Time of Flight Mass Spectrometer Developer & Manufacturer

Department: Hardware & Software Customer Support

GUI: “GUI.png” & “EventLog.png”

Sample Raw Data:

“InputDataForm.png”, “InputDataForm_Top.png”, “InputDataForm_Middle.png”, & “InputDataForm_Bottom.png”.  Each member of the service department will fill out this form and email it to the service department manager each week.  Every possible task that a member of the service department team could spend time doing have been broken into 13 categories.  Each category is further sub-divided into 3 to 14 subcategories.  The employee is expected to account for their time during the course of a week.

Sample Output:

“SampleOutput_EntireTeam_AllCategories.png” Provides an example report when the data base was queried for the total man hours spent by the entire service department on each of the 13 categories.

“SampleOutput_EntireTeam_SingleCategory.png” Provides an example report when the data base was queried for the total man hours spent by the entire service department on a single category thus the corresponding subcategories are shown.

“SampleOutput_IndividualTeamMember_AllCategories.png” Provides an example report when the data base was queried for the total man hours spent by an individual member of the service department on each of the 13 categories.

“SampleOutput_ IndividualTeamMember _SingleCategory.png” Provides an example report when the data base was queried for the total man hours spent by an individual member of the service department on a single category thus the corresponding subcategories are shown.

Application Description:

The workflow for utilizing this application essentially has three stages…

1) Team members input their allocated man hours into the data input file, which is a protected excel spread sheet.  At the end of every week, they email the spread sheet to the service department manager.

2) The data is uploaded into the database using the GUI which allows for the selection of multiple files at a time.

3) Using the GUI the manager can define how the database is queried.  Next, a new excel spreadsheet is generated and the data is plotted and tabulated.

The major challenge I encountered with this project was the timeline.  My manager wanted a functioning beta version in 11 days.  I was able to deliver, but I had to cut some corners and keep things very basic.  The final version of this was written in Python & SQL which was much more powerful with far more features.  See the TimeTrax_06-24-2017 repository for the final version of this application.

