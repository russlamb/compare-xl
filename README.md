# Automatically Compare Two Databases or Excel Files 

Thursday, February 08, 2018

5:48 PM

## Overview 

Often when making changes to a stored procedure or SQL object, testing if the changes have downstream impact can be difficult.

 

This tool automatically runs SQL on two servers, stores the results in Excel files, then compares the Excel files and stores differences in a spreadsheet.

 

### Key benefits include

1.  Automatically saves a hard copy of the results of your sql query from each environment, reducing context switching. This means better ease of testing and reduces human error.

1.  Lists all differences for you and saves in a file, meaning fewer mistakes from manual error

1.  Quick comparisons between production and Development or UAT systems

2.  Compare any excel files quickly and automatically, meaning less developer time wasted doing excel comparisons
 

## Setup

These instructions were made using the Windows 7 Operating system. Other operating systems should work similarly since cx\_freeze is cross platform.

Prepare Python Environment
--------------------------

My first attempt to build was done in the system-wide install of python, but this caused cx\_freeze to crash.

To fix this, I created a virtualenv for the project and only installed what was needed.

The necessary packages are found in the requirements.txt file. Install using pip and you should be good to go.

### Make executable
---------------

To deploy as an executable, build the project by executing the following command:

python setup.py build

This will create a directory called “build” in the local directory with another subdirectory containing your executable. Zip that folder and deploy to your desired location.

### After building
--------------

1.  Unzip the folder on your workstation.

1.  Open a command line

2.  Navigate to the folder you unzipped

1.  Run the following command: compare\_data.exe -h

This should display the help text for the executable. If you get an error instead, see Troubleshooting.



## Troubleshooting

If you get an error like this then you need to install the Microsoft Visual C++ package.

![Machine generated alternative text: ‘1compare\_data.exe-System Error
Q The program can’t start because api-ms-win-crt-stdio-l1-1-OdIl is
‘ missing from your computer. Try reinstalling the program to fix this
problem.](/media/image1.png){width="4.395833333333333in" height="1.65625in"}

 

See [this link](http://www.thewindowsclub.com/api-ms-win-crt-runtime-l1-1-0-dll-is-missing) for more information.

 

#### Download links

Depending on your machine, you may need one or the other

[32 bit download](http://www.microsoft.com/en-gb/download/details.aspx?id=5555)

[64 bit download](http://www.microsoft.com/en-us/download/details.aspx?id=14632)

 

## How To Use

Some default parameters can be configured using the config.ini file contained in the folder.

 

### Run a SQL command on two databases and Compare

 

1.  Open a command line

1.  Navigate to the folder where you unzipped the files

1.  Run the command: **compare\_data.exe -sql "select AssetID from Asset"**

Substitute "**select AssetID from Asset**" with your own SQL statement. Make sure to enclose the statement in quotes.

 

This will cause the executable to do the following:

1.  run the sql you passed in against Prod and save the results in Left.xlsx

1.  Run the sql against QA and save results in Right.xlsx

1.  Compare the two excel files and save differences in Differences.xlsx

>  



**Compare two existing Excel files**

After running, the example above you can try this example

 

1.  Open Command line

1.  Navigate to the folder

2.  Run the command: **compare\_data.exe -compare -left Left.xlsx -right Right.xlsx**

 

To compare your own files, substitute **Left.xlsx** and **Right.xlsx** with of your own files.

 

This will cause the executable to:

1.  Compare the files you passed in and save differences in Differences.xlsx

 

### Examples

#### Example using defaults

Compare\_data.exe -sql "select AssetID from Asset"

 

#### Example Specifying overrides

If I want to change the file names for left and right, I can do the following:

 

Compare\_data.exe -sql "select AssetID from Asset" -left "prod.xlsx" -right "test.xlsx"

 

This will store the production data in prod.xlsx and QA data in test.xlsx

 

#### List of command line flags and their defaults

By default, the tool uses the following parameters that can be overridden on the command line:

 | Argument | Default | Description |
|-----------|--------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| -left | ./Left.xlsx | First file for comparison |
| -right | ./Right.xlsx | Second file for comparison |
| -diff | ./Differences.xlsx | Differences between files |
| -db_left | Prod database | Connection string for a database |
| -db_right | QA database | Connection string for a database |
| -index | None | (optional) An integer indicating   which column to use as an index when comparing files.  This will line up the two files so that the   index values match when comparing differences.  If index value is not unique, it will use   the last row with each given index value.    Non-intersecting rows are excluded from the comparison. |
| -sql | None | A SQL statement  |
| -sheet | Output | Any name you want for the sheet |
| -compare | None | Use this flag to compare existing   excel files.  No SQL will be run.   |