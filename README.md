# puller.ps1 ***NEW AND IMPROVED***
The puller powershell script offers a better user interface. The code no longer needs to be edited. Multiple ships can be done at once; use a comma to delimit the ships in the text box. Only the email addresses will be scraped, if the checkbox is left checked. If the checkbox is unchecked, all the information will be scraped, including email address, last name, first name, middle initial, modified time stamp, and distinguished name.

Internet explorer window is also hidden, the $ie.visible = $true line (295) can be uncommented for troubleshooting. When you run the script, you will need to choose your CAC cert and type in your pin. The DOD warning banner needs to be accepted before providing the ship hull numbers. The connection may timeout if left too long, but this was not tested.

The script will write the output for each ship into a file named <ships-hull-number>.txt. For example, lhd7 output will be written to lhd7.txt. The output files will be written to the directory that the script is running from.

To run, make sure your powershell execution policy is not restricted. To verify, run:

Get-ExecutionPolicy

To allow this scripts (and scripts in general to run) use the following command:

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Then you can run the script:

.\puller.ps1

# puller.vbs ***DEPRECATED***
***This script is included for reference only***
The puller application will interact with an open Internet Explorer window to click through an email list and write the emails on each page to a file. The code only works on a specific website. Improvements still need to be made.

To use, verify the output file in the puller.vbs file is correct then run the following in a command prompt:

cscript.exe puller.vbs

# lastn_puller.vbs ***WORKING VB SCRIPT***
The lastn_puller application will pull emails for lists that are longer than 500. The initial A and proper selections need to be made on the website, then run it the same way as the puller application:

cscript.exe lastn_puller.vbs

# auto_keys.vbs
The auto_keys application will write the contents of a file into whatever window is currently active, so whichever window you click into after running the application. This is sometimes buggy, especially when writing into a virtual machine. It helps if you hover your mouse over the window and keep the mouse active over the window (i.e. move the mouse in small circles over the desired window). You have to change the "read from" file within the code. To run, use the following:

wscript.exe auto_keys.vbs

Still need to rewrite the auto_keys for powershell
