@ECHO OFF
ECHO [+] Remove TEMP...
RD /S /Q _TEMP_

FOR %%T IN (educate finance forms layout misc officorr offimisc personal presnt) DO CALL cretwork-ml.bat %%T %1

ECHO [+] Remove OOOP-ml-templates-%1.oxt...
DEL OOOP-ml-templates-%1.oxt

ECHO [+] Creating empty OOOP-ml-templates-%1.oxt...
COPY MultiLanguageTemplatePackage.oxt OOOP-ml-templates-%1.oxt 

ECHO [+] Compressing to OOOP-ml-templates-%1.oxt...
7Z.EXE A -MX=9 -R OOOP-ml-templates-%1.oxt .\_TEMP_\*

ECHO [+] Compressing to OOOP-ml-templates-all.oxt...
7Z.EXE A -MX=9 -R OOOP-ml-templates-all.oxt .\_TEMP_\*

ECHO [+] Remove TEMP...
RD /S /Q _TEMP_

ECHO [+] Finished, exiting and go home...
ECHO.