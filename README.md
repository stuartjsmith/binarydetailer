# Binary Detailer
Binary Detailer is a method of outputting binary details such as net framework, 64-bit compatibility, version etc. Given a directory name, it will iterate dll and exe files and output the results to a csv file.
___

### Usage

1. Build the project to produce the BinaryDetailer.exe
2. Open a Windows terminal
3. Call the .exe with file path as an argument
___

### Example

![Screenshot 1](data/screenshots/CMD-Arguments.jpg "CMD Arguments")

![Screenshot 2](data/screenshots/complete.jpg "Complete")

![Screenshot 3](data/screenshots/Completed-csv.jpg "Completed CSV File")
___

### Arguments

'config' - This requires an XML passed after the config argument. This will compare the binaries against the XML and group common binaries together. See "GroupingConfigExample.xml"
'doc' - This will create a word document table of the data. This requires a Microsoft Office install on the machine executing.

Example argument: BinaryDetailer.exe "C:\Program Files\dotnet\sdk\6.0.400" config "C:\BinaryDetailer\GroupingConfigExample.xml" doc

## Contribute

Congratulations! You’re up and running. Now you can begin using and contributing your fixes and new features to the project.
