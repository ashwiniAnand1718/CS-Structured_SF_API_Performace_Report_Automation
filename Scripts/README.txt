This Document will helps to run the python automation script for the  SF API performance Report 

Pre-requsitits 
1. Python installed on the system
2. Module that needs to be installed : openpyxl, os , pillow.
3. Folder format should present in the following way (Mandatory step) 
      * Download_result : Where the raw report file will be downlaoded from jenkins.
	  Ex: "C:\\Users\\sa\\PycharmProjects\\PT_report_formatting\\Download_result\\Windows" & "C:\\Users\\sa\\PycharmProjects\\PT_report_formatting\\Download_result\\Linux"
	  * Template        : Which contains blank template for the reporting.  EX: "C:\\Users\\sa\\PycharmProjects\\PT_report_formatting\\New_template"
      * Formatted_report: Which will be the final formatted performance report with release version updated.  	Ex:"C:\\Users\\sa\\PycharmProjects\\PT_report_formatting\\Formatted_report"	  


Scripts Name :
    1. Download_Raw_file: This script will download the raw report from jenkins 
    2. Formatting_Report_Script : This script will format the copy the data from the raw report to final report template.
    3. Release_version : Will update the release version into all the final report & appends the file name with the release version 
	  

Workflow: 

Step 1: First we need to run " Download_Raw_file.py" which will  download and save the raw report from Jenkins url to their respective folder (as menetioned in pre-requsits 3 step)
     So all windows report will be downloaded and saved in Windows spcefic folder & Linux reports will be downlaoded and saved in spcefic Linux folder.

Step 2: Run " Formatting_Report_Script" Which will first copy all the empty report template from "template" folder to " Formatted_report" folder. Then it copies all datas from raw reports and pastes it into the formatter report. At the end to also formats all the data with "All boarder" and does the "center"  alignment.


Step 3: Run "Release_version" , this script will ask you to enter/input  the release version that needs to be updated in all the report. Once we eneter the release version, it will update the version into all the report & it will also  append it to the report file names accordingly.


