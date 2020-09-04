# Script-OutlookClient-Search-email

Parameters/flags:

-email_address \<string\> : set the targeted mailbox (name show in outlook)  
-mailbox_folder \<string\> : set the targeted folder with a 1st lvl folder name (i.e, "inbox")  
-skip \<string\> : skip folders containing this string  
-keyword \<string\> : search emails containing this string. Default : search in email's subject   
-output_path \<string\> : set a path for the output. Default path is the current location  
-output_filename \<string\> : set the output filename (.csv) containing the results. Required if need an output file  
-startDate \<date\> : Set the start date of the search using international date format (DD/MM/YYYY HH:mm:ss)  
-endDate \<date\> : Set the end date of the search using international date format (DD/MM/YYYY HH:mm:ss). Default value is the current datetime.  
[-output_full_option] : Full output result. Required the -ouput_filename flag enabled  
[-export_emails_html] : Export emails results in html file  
[-deep] : Deep search using html body  

Information:

Script that uses Outlook client for counting, searching and exporting emails.  
Outlook client should be closed in order to have the best results  

Examples:

Count the number of mails within a mailbox:  
.\script.psl -email_address "example@exemple.com"

Deep search and email export:  
.\script.psl -email_address "example@exemple.com" -mailbox_folder "Inbox" -skip 9 -keyword "toto" -output_path folder_example -output_filename output.csv -export_emails_html -deep
