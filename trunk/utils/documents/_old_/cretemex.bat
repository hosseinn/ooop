@ECHO OFF
ECHO [+] Remove TEMP...
RD /S /Q _TEMP_
FOR %%T IN (educate finance forms layout misc officorr offimisc personal presnt) DO CALL cretwork.bat %%T %1

ECHO [+] Adding licensing information...
MD .\_TEMP_\licenses
XCOPY ..\..\documents\license\extension\license_*.txt .\_TEMP_\licenses\

ECHO [+] Adding readme information...
MD .\_TEMP_\readmes
XCOPY ..\..\documents\license\extension\readme_*.txt .\_TEMP_\readmes\

ECHO [+] Adding description information...
XCOPY ..\..\documents\license\extension\example\description.xml .\_TEMP_\

ECHO [+] Remove OOOP-templates-%1.oxt...
DEL OOOP-templates-%1.oxt

ECHO [+] Creating empty OOOP-templates-%1.oxt...
COPY TemplatePackage.oxt OOOP-templates-%1.oxt 

ECHO [+] Compressing to OOOP-templates-%1.oxt...
7Z.EXE A -MX=9 -R OOOP-templates-%1.oxt .\_TEMP_\*

ECHO [+] Remove TEMP...
RD /S /Q _TEMP_

ECHO [+] Finished, exiting and go home...
ECHO.