@ECHO OFF
ECHO [+] Processing "%1" folder, [%2] language...
MD .\_TEMP_\template\%1
XCOPY ..\..\extras\source\premium\templates\%1\lang\%2\*.* .\_TEMP_\template\%1\  /EXCLUDE:cretwork.exc
