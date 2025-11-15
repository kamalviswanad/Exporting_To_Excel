# Exporting_To_Excel

This code runs a command on a remotly controlled server and gets information related to Processes. And then the information is exported into an Excel sheet. 

You can run it on one or more servers.

To use the function fill out the server information and fail path in "FinalOutput" function in the MainFile.psm1 file. You can save the required credentials in Windows Credential Manager so you don't need to fill in the information everytime you run the script. You can just call the function by typing "FinalOutput" where ever you want by importing it in your scripts.
