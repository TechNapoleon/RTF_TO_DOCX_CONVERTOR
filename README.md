# RTF_TO_DOCX_CONVERTOR


Preface:

This project created due to the requirement to output all the documents of the system I work with in format of DOCX instead of RTF format. However, because of the difficulty in changing the way the system works, I wrote this script.


Overview of the program:

The script is always works in the background and checks every one second whether the flag (txt file) is exists. Once the flag file does exist, the script walks throe folders where the documents are created (The paths are predefined in the code), its converts it to DOCX format and delete the original RTF file. 
The flag method created this way to be able to add to the system flow something like this code – 
Create file=flag.txt
While flag.txt exist:
	sleep(1)
This way you will be able to let the system to continue the process of the file after the creation and the script will not take many resources.
However, it possible to rewrite the script in anyway you required to. I have previously used the script with argument method of path to the RTF without any issue.
If this script needs to be run automatically when the server is started, please read the notes section!!!


Installation:

1.	Download the git repo.
2.	You need an Office installation on the server/pc.
3.	Install library pywin32 (python -m pip install pywin32)

Notes:

If required to run the script in Windows environment as batch job (Task Scheduler or Service) there will be a problem. Microsoft doesn’t allow by default to run the Office in service mode. Meaning you will not able to run pywin32 in the batch, you can run it just with your own user while you are logon. 
Because when you execute pretty much anything in the batch, you execute it with system user which have no environment to open the Word or any other Office application.
To overcome this issue, you can create the environment by creating these paths – 
Windows Server x64
C:\Windows\SysWOW64\config\systemprofile\Desktop
Windows Server x86
C:\Windows\System32\config\systemprofile\Desktop

If it doesn’t work for you anyway, I recommend you read this links – 
Microsoft explanation - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/257757

Forum about this issue – https://social.msdn.microsoft.com/Forums/en-US/b81a3c4e-62db-488b-af06-44421818ef91/excel-2007-automation-on-top-of-a-windows-server-2008-x64?forum=innovateonoffice

Please contact me if there are any bugs or in case you have any questions.
