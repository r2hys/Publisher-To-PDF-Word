# Publisher-To-PDF-Word
This is a script that will bulk convert publisher files to either .PDF or .DOCX

 

Configure Script 

Change the Root Folder variable in the script to your File Path  

You can use an IDE of your choice (I prefer Visual Studio) or use notepad for this  

NOTE: DO NOT USE ISE, THIS SCRIPT WILL NOT WORK WITH IT 

 

File / Folder Setup 

Open word / Publisher before running the script and accept any T&Cs (You can close them after this) 


Place the PowerShell script and the CMD file in a folder on the machine, preferably in the C Drive but I believe anywhere works 

 


Running The Script 

Double click the .CMD file to begin the script 


Running a dry run is optional, the more files the longer it takes to complete (This is good if you want a rough idea of how many .PUB files exist in your file path). 

If you choose Dry Run, you will have to re-run the CMD file to start converting. 

NOTE: THIS SCRIPT WILL NOT DELETE THE ORIGINAL PUB FILE, IT WILL DUPLICATE IT AS YOUR CHOSEN EXTENSION 

From Here the script will run and create your new files, they will live in the same folder as the .PUB’s it converted. Remember, the more .PUB files the longer this will take.  




Log Files 

Log Files can be found in C:\PublisherLogs 

You can edit this in the script if you wish 

 

Important Info 

I have noticed that the Word conversion can be a hit or miss with some files, and if you do choose the word route, it will automatically create a PDF too. (I can’t get it to create a word doc without this) 

If you want an easy and accurate conversion, please choose PDF (will need an adobe licence to edit these though)

 

It is still relatively early stages, however bits I have tested work fine.  

