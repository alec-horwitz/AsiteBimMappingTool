CONTENTS OF THIS FILE
---------------------
-BimAsiteMappingTool.exe

-Data/settings.txt

-Data/ReadMe.txt

-Data/BimAsiteStatusTrans.txt

-Archive/




 TABLE OF CONTENTS
---------------------
-Introduction

-Configuration

-Troubleshooting

-Support




   INTRODUCTION
---------------------
This program is made to help identify the answer to three important 
questions about the differences between the BIM 360 and Asite databases 
for the prequalification process. The three questions this program helps 
answer are as Follows:

-Are companies in one database missing from the other and if so which 
 database is missing which companies?

-What are differences in how company names are spelled between the two 
 databases, if any?

-What are differences in the company statuses, for the prequalification 
 process, between databases?

This program produces a Tab Separated Value(.xlsx) file with no formatting 
but with all the above mentioned information. This .xlsx file can be opened 
with Microsoft Excel.

In order for this program to produce this data properly you must follow 
the instructions provided in the CONFIGURATION section.




   CONFIGURATION
---------------------
This section of the ReadMe is a detailed description of how this 
program functions. It explains how to give the program an input 
that it expects and why. This section should be able to explain any 
problems you might run into and so it also happens to be the longest 
section. As such for users that are already comfortable with how the 
program works on a whole refer to the TRUBLESHOOTING section. We will 
start with the settings.txt and then move onto the other inputs the 
program reads in.


#####About the settings.txt
The settings.txt file must be in the same directory as this program. 
This program is very particular about how this file is set up. Do not
change the name of this file and do not add lines that aren't required. 
If you do, it might give you an error message (or worse, crash without 
any error message). It must have 21 lines(including the 3 blank lines) 
in the given order that read as follows:
[Asite Input]
ASITE_INPUT_DIR: 
Asite_Header_Row_Num: 0
Asite_ID_Col_Num: 0
Asite_Company_Col_Num: 0
Asite_Status_Col_Num: 0

[BIM Input]
BIM_INPUT_DIR: 
BIM_Header_Row_Num: 0
BIM_ID_Col_Num: 0
BIM_Company_Col_Num: 0
BIM_Status_Col_Num: 0

[Status Translation File]
BIM_TO_ASITE_STATUS_TRANSLATIONS_PATH: Data/BimAsiteStatusTrans.txt
Translation_Delimiter: , 

[Output]
OUTPUT_PATH: preqMap.xlsx
DUMP_FOLDER: Archive/

For each line with a ":" in it anything after the ":" can be altered to your 
liking so long as you edit it correctly. Each line in the above example that 
have file paths with files at the end of them are there as examples of the 
correct file types to use for those lines. If you change the file paths in the 
settings.txt make sure that file types match the example above for the appropriate 
lines. There are several different sets of rules you must follow to edit these 
lines correctly depending on what is just before the ":". These rules are as 
follows:

* If the line has a "_DIR:" then you must give it a path to a folder that contains
 a file of the appropriate file type. There are two ways you can enter a path and 
 the program does not care which you use as long as you are consistent. Being 
 inconsistent with which method you use to describe a file path may produce unforeseen 
 consequences and will defiantly cause the program to not function correctly. The
 two ways you can enter a path are as follows:

  * You can enter a relative(to the directory the program is in) file path 
   such as "relative/file/path/". This tells the program to look for a file
   with in a sub directory of the folder that the program lives in. If you 
   give it no path after "_DIR:" the program will look for the file in the
   exact same directory as where the program lives. For example:

    * If you were to type "Data/" after the line "ASITE_INPUT_DIR:" in the form 
    "ASITE_INPUT_DIR: Data/" the program will look for a .xls file containing 
    prequalification data in a folder called "Data" in the same directory that 
    the program is stored in. 

    * If you were to type nothing after the line "ASITE_INPUT_DIR:" the program 
    will look for a .xls file containing prequalification data in same folder 
    that the program is stored in.

  * You can enter an absolute file path such as "N:/whatever/path/you/want/".
   If you are an experienced Windows or Linux user this should be the
   type of directory scheme you are used to seeing(In Windows they use 
   '\' instead of '/' but this program doesn't care which you use). For example
   if you typed "C:/Desktop/" after "BIM_INPUT_DIR:" in the form 
   "BIM_INPUT_DIR: C:/Desktop/" the program will look for a .csv file containing 
   prequalification data on the Desktop folder on the C drive.

* There are also two ways you can have this program look for a file:

  * You can specify just the path to the file. If you do this the program will 
   read in the first file, in the given file path, it sees with the correct file 
   extension. As such, you should make sure that only one file with the correct 
   file extension exists in the directory you give it. For example if you want 
   to tell the program to find the BIM prequalification data in "C:/Downloads/" 
   you simply type that after the "BIM_INPUT_DIR:" in the form of
   "BIM_INPUT_DIR: C:/Downloads/". This will tell the program to look for a .csv
   file in the Downloads folder on the C drive.

  * You can specify an explicit file name. If you do this the program will look
   for a file with the exact name you provide it in the directory you provide 
   it. For example if you want to tell the program to find the Asite
   prequalification data in "N:/PROGRAMS/" and the file name is "Form Listing.xls"
   would type "N:/PROGRAMS/Form Listing.xls" after "ASITE_INPUT_DIR:" in the form
   of "ASITE_INPUT_DIR: N:/PROGRAMS/Form Listing.xls". 

* If the line has a "_PATH:" you can treat it exactly like a line with "_DIR:" at the end of
 it except you must give it a filename at the end of the file path you enter. For example
 if there is a line "BIM_TO_ASITE_STATUS_TRANSLATIONS_PATH:" you may enter a path after it 
 in the form of "BIM2AsiteStatusTranslations/BimAsiteStatusTrans.txt" or even 
 "N:/whatever/path/you/want/fileName.txt" but you may not enter a path without a filename 
 and proper extension at the end of it like "relative/file/path/" or "N:/some/path/". If you 
 give it a file path like "relative/file/path/" or "N:/some/path/" you will get an error 
 message saying that it can't find the file.

  * If the line is "OUTPUT_PATH:" you must treat it just like a normal "_PATH" line but you can 
   make the filename whatever you'd like and the program will append the date and military time 
   to the end of the name right before the file extension. For example if the military time was 
   4:33:25 seconds, the date was 7/19/2016 and you want the output file to be generated in 
   "C:/Documents" and you want its name to be "usefulData.xlsx" you would type 
   "C:/Documents/usefulData.xlsx" after "OUTPUT_PATH:" in the form 
   "OUTPUT_PATH:C:/Documents/usefulData.xlsx". If you then ran the program a file called 
   "usefulData-160719-163325.xlsx" would get generated in the Documents folder on the C drive. If 
   another .xlsx file already exists in that folder it will be moved to the path given after the line 
   "DUMP_FOLDER:" in settings.txt. As such make sure that there aren't any other .xlsx files in that 
   folder besides the last output file that was generated. 

* If the line is "DUMP_FOLDER:" you can treat it exactly like a line with "_DIR:" at
 the end of it except you cannot give it a filename at the end of the file path you
 enter. For example for "DUMP_FOLDER:" you may enter a path after it in the form of 
 "relative/file/path/" or even "N:/whatever/path/you/want/" but you may not enter a 
 path with a filename at the end of it like 
 "relative/file/path/filename.extension" or "N:/some/path/filename.extension". If 
 you give it a file path like "relative/file/path/filename.extension" or 
 "N:/some/path/filename.exstension" you will get an error message saying that it 
 can't find the file.

* If the line has "_Header_Row_Num:" you must specify what row, in the spreadsheet the program 
 should start reading in the data from. For example if there's a header on row 1 for 
 the title of the document and a header on row 2 for the names of each row and the first
 of the data you want read in is on row 3 you would type "2" after "_Row_Num:" in the form
 "_Row_Num: 2". If you enter the number 0 for these lines this program will attempt to find the
 the header rows automatically.

* If the line has "_Col_Num:" you must specify what column, in the spreadsheet the program 
 should be reading in data from. For example if in the BIM prequalification listing the 
 Company IDs were in column 4 then you would type "4" after the line "BIM_ID_Col_Num:"
 in the form "BIM_ID_Col_Num: 4". If you enter the number 0 for these lines this program will 
 attempt to find the the column numbers automatically.

* If the line is "Translation_Delimiter:" enter what you want the delimiter to be between
 the Bim statuses if there are more than one Bim statuses that translate to just one Asite
 status. For example if in your translation .txt(this will be covered later on in detail) 
 file you have a line like:
 "Pending Approval: Requested, Incomplete, To be Contacted, Received, Evaluated, Reviewed" 
 you must indicate that you want the delimiter to be "," by typing it after the line 
 "Translation_Delimiter:" in the form "Translation_Delimiter: ," in the settings.txt file.
 If you decided that you wanted to use ","s in the status names you could choose a different
 delimiter so that those statuses are not read as two separate statuses. For example you have 
 a line translation .txt file like:
 "Pending Approval: Requested, Ralph, check this out, Received, Evaluated, Reviewed"
 and you wanted "Ralph, check this out" to be read as one status you would first pick a new 
 delimiter, say "/", and replace all the ","s in the translation .txt file with "/" except 
 for the ","s that you want to be part of a status so the sample line would end up like 
 this: "Pending Approval: Requested/ Ralph, check this out/ Received/ Evaluated/ Reviewed"
 You would the update the settings.txt file
 to reflect that by typing "/" instead of "," after the line "Translation_Delimiter:" in the 
 form "Translation_Delimiter: /". The program would the read "Ralph, check this out" as one
 status without a problem.


##### Advanced Program Functions for File Handling 
If the path given after "DUMP_FOLDER:" is the same as the path given for any 
or all "_DIR:" lines, the data of the file found in the path(s) will be read in without any changes 
to the filename(s).

If the path given after "DUMP_FOLDER:" is not the same as the path given for any or all "_DIR:" 
lines, the data of the file found in the path(s) will be read in and the filename(s) will be 
changed with the date and military time appended in the format 
"assumedNameOfDataBaseDir-YearMonthDay-HourMinuteSecound"
for example if the military time was 13:04:23 seconds, the date was 7/19/2016, the path given after 
"DUMP_FOLDER:" was "archive/" and the path given after "ASITE_INPUT_DIR:" was "data/" then the .xls 
file found in the data folder in the program's directory would be renamed "AsiteDir-160719-130423.xls".

If the path given after "DUMP_FOLDER:" is the same as the path given for "OUTPUT_PATH:" then the program
will ignore it and generate a new output file as it normally would with a more current time stamp in the 
name. For example if the military time was 21:14:47 seconds, the date was 7/19/2016, the path given after 
"DUMP_FOLDER:" was "archive/", the path given after "OUTPUT_PATH:" was "archive/outputFilename.xlsx" and 
there was already a file in the archive folder called "outputFilename-071916-130423.xlsx", the program would
simply generate a new file called "outputFilename-160719-211447.xlsx" along side it in the same folder 
without consequence.


##### About the translation .txt 
The purpose of the translation .txt is so that this program can match a BIM 
prequalification status to its Asite equivalent. Regardless of whether there is a 
difference in statuses between the two databases the still must be a the 
translation .txt file for the program to read, otherwise you will get an error 
message. The translation .txt can be in any directory as long as you declare its 
location in the settings.txt. The translation .txt can also be named anything as 
long as its name has the extension .txt at the end of it and you declare its name in the 
settings.txt. You can add or delete as many lines to this file(each line containing a 
different status translation) as you wish as long as they are written in the correct format. 
Making edits to this file incorrectly will not cause the program to produce a error 
message except if there is no data in it at all. Editing this file incorrectly has the 
potential to crash the program without an error message and if not the program will simply 
produce a file with incorrect information about the statuses. As such it is important to 
edit this file correctly. For example here are the status translations as of July 18, 2016:

Approval on Hold: Hold
Rejected - DO NOT USE: Rejected
Declined to Submit: Declined
Unresponsive: Unresponsive
Owner-Preferred (Not for CM-at-Risk): Owner
Approved: Martin
Pending Approval: Requested, Incomplete, To be Contacted, Received, Evaluated, Reviewed

Notice that the Asite statuses are on the left side of the ":" and only unique key words
of the BIM statuses are on the right side of the ":". This is the format in which status
translations must be written in the translation .txt.

There Must only be one Asite status per line but there may be multiple BIM statuses per 
line.

If there are multiple BIM statuses that translate to one Asite statuses you may list key 
words or word combinations that uniquely identify the BIM status after the ":" using what 
ever delimiter you want so long as you declare the delimiter in the settings.txt. If you 
do not declare the delimiter that you use the program will read the entire list as one 
status name and has the potential to crash the program or produce incorrect data. For
example if the delimiter is declared as "," in the settings.txt then you could write
a line like the following to indicate multiple BIM statuses translate to one Asite status:

Pending Approval: Requested, Incomplete, To be Contacted, Received, Evaluated, Reviewed

Note that there are no spaces after each key word(s). If you were to put a space after
a key word(s) like "To be Contacted ," the program will read the key word in as including 
the space. To ensure that the program produces accurate data do not leave spaces after a
key word(s).

If there are multiple Asite statuses that translate to one BIM status there must be 
multiple lines(one for each Asite status) with the key word or word combination that 
uniquely identifies the BIM status after the ":". For example if "Requested" and 
"Incomplete" were both Asite statuses and they both equated to a "Pending Approval"
status on BIM and "Pending" was a unique key word for that BIM status then you would 
type the following lines in the translation .txt:

Requested: Pending
Incomplete: Pending

Note that there are no spaces before or after the Asite statuses. If you put a space
before or after a Asite status like " Incomplete :" the program will read the 
status in as including the spaces. To ensure that the program produces accurate data 
do not leave spaces before or after the Asite statuses.




   TRUBLESHOOTING 
---------------------
This section of the ReadMe focuses on just the error messages and reasons why 
the program would appear to not run produce no error messages. Here you will find 
a list of error messages followed by solves to them. If the program has no error 
message and seems to run and do nothing there are possible solves for that at the 
end of this section. This section is meant as a quick on the fly refresher on what 
problems you might be running into. This section does not provide any new information 
or solutions that you couldn't get from the CONFIGURATION section and is by no means 
a substitute for it. To ensure proper resolution to problems you might be having 
with this program please be sure to read the CONFIGURATION section. If you have 
read this section and the CONFIGURATION section and are still unable to have the 
program run properly refer to the SUPPORT section.


"ERROR: Cannot find the settings.txt file.
 This program cannot run properly without it.
 Enter any key to EXIT program..."

If you get this error either settings.txt is missing from the file directory
that you are keeping this program in or the settings.txt was renamed and the 
program no longer recognizes it.


"ERROR: \\NLS-FS01\some\path is not a valid path.
Closing all File Explorer windows and try again.		    
Enter any key to EXIT program..."

If you get this error close all File explorer windows and start over.
Do not use short cuts to get back to the N drive. Go to the N drive by 
Going to the "This PC" folder and then to the N drive. The N drive has
an alternitive directory path that can not be read by this program. This 
error message only comes up if you use absolute file paths or if you are 
using this programs drag and drop featcher. 


"ERROR: Could not find the NameOfColumn column in one of the input files!!!
This program cannot run properly without it.
To fix this manually enter the appropriate column number in the settings.txt
Consult the ReadMe.txt if you are not sure of how to properly modify the file.
Enter any key to EXIT program..."

If you get this Error then you used then you set this program to automatically
find the column numbers of the data that needs to be read in. If this is true
then the programs attempt to automatically find these numbers failed. Change
the entre after the "_col:" line in the settings.txt to reflect the actual 
column where the data can be found. Please check the CONFIGURATION section in 
this read me if you are unsure of how to do this correctly.


"ERROR: Cannot find the following A_Line_From_Settings_File: some/file/path
 This program cannot run properly without it.
 Please check that the file path is entered correctly in the settings.txt file.
 Consult the ReadMe.txt if you are not sure of how to properly modify the file.
 Enter any key to EXIT program..."

If you get this error the path, filename, and/or extension was entered wrong in 
settings.txt or the file or path does not exist. Make sure that the path, filename, 
and extension is indeed entered correctly in the settings.txt and the file and path
does in fact exist.


"ERROR: some/relative/path/to/a/file.extension is empty!!!
 This program cannot run properly without it.
 Please check that required information is properly entered in this file.
 Consult the ReadMe.txt if you are not sure of how to properly modify the file.
 Enter any key to EXIT program..."

If you get this error either the file has no text in it. Make sure the indicated
file has at least some number of line of text in the appropriate format as explained
in the CONFIGURATION section of this ReadMe. This error should only pop up if the 
settings.txt or the translation .txt is blank.


"ERROR: the line something_somthing: is missing or not where it's supposed to be in settings.txt!!!
 This program cannot run properly without it.
 Please check that required information is properly entered in this file.
 Consult the ReadMe.txt if you are not sure of how to properly modify the file.
 Enter any key to EXIT program..."

If you get this error the content of settings.txt is either missing lines that the 
program expects, there are extra lines the program doesn't expect, or the lines 
are out of the expected sequence. Make sure settings.txt's configuration matches 
the description in the CONFIGURATION section.

"ERROR: The line "OUTPUT_PATH:" in settings has no file path to write to.
 This program cannot run properly without it.
 Please check that the file path is entered correctly in the settings.txt file.
 Consult the ReadMe.txt if you are not sure of how to properly modify the file.
 Enter any key to EXIT program..."

If you get this error then you did not put a filename or a path with a filename at 
the end of it for the line labeled "OUTPUT_PATH:" in the settings.txt. You can put in
path you'd like as long as it exists and any filename as long as it has .xlsx at the 
end of it but you must put something so the program knows where to put the file and 
what to name it.



##### IF SOMETHING WENT WRONG AND YOU GOT NO ERROR MESSAGE 

If you get no error message and the program produces nothing then most the likely
reason is that there is something wrong with the file types specified in the 
settings.txt. Make sure that in the settings.txt:
 
-The line "ASITE_INPUT_DIR:" has a path to an .xls file that exists with the correct 
 data in the rows specified.

-The line "BIM_INPUT_DIR:" has a path to a .csv file that exists with the correct data 
 in the rows specified.

-The line "BIM_TO_ASITE_STATUS_TRANSLATIONS_PATH:" has a path to a .txt file that exists 
 with the correct data and formatting.

If you don't it will result in a fatal error with no error message. In other words the 
program will just crash quietly with no explanation.


If the program produces an output file that seems to contain seemingly random 
combinations of letters, numbers, and symbols then this is probably do to the output
file described in the settings.txt having something other than .xlsx at the end of it.
Be sure to specify the correct file type.


If the program completes the operation and produces a file with incorrect data about the 
statuses then the data in the translation .txt file is wrong or formatted incorrectly.


If you have a file that has x number of rows and in the settings.txt for a line labeled 
"_row_num:" for the appropriate file has the number y at the end of it and y>x the program 
will crash. Make sure that the rows you enter in the settings.txt exists in the file the 
program is reading in. 

If you have a file that has x number of columns and in the settings.txt for a line labeled 
"_col_num:" for the appropriate file has the number y at the end of it and y>x the program 
will crash. Make sure that the columns you enter in the settings.txt exists in the file the 
program is reading in. 


If you use both relative paths and absolute paths in the settings.txt this may be the
cause of a variety of issues. Make sure this is not the case.


If you have an old output file opened in excel when you run the program this will also 
crash the program with no error message.



   SUPPORT
---------------------
If user are having issues and this document has not helped you may contact me the creator 
of the program at:

horwitz_alec@wheatoncollege.edu

I will try my best to resolve the issue as soon as possible but I am a full time student
so it might take a while to get back to you.
