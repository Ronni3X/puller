# puller
The puller application will interact with an open Internet Explorer window to click through an email list and write the emails on each page to a file. The code only works on a specific website. Improvements still need to be made.

To use, verify the output file in the puller.vbs file is correct then run the following in a command prompt:

cscript.exe puller.vbs

# lastn_puller
The lastn_puller application will pull emails for lists that are longer than 500. The initial A and proper selections need to be made on the website, then run it the same way as the puller application:

cscript.exe lastn_puller.vbs

# auto_keys
The auto_keys application will write the contents of a file into whatever window is currently active, so whichever window you click into after running the application. This is sometimes buggy, especially when writing into a virtual machine. It helps if you hover your mouse over the window and keep the mouse active over the window (i.e. move the mouse in small circles over the desired window). You have to change the "read from" file within the code. To run, use the following:

wscript.exe auto_keys.vbs

Eventually all of these applications will be combined into one and will be differentiated by command line arguments.
