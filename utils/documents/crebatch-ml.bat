@ECHO OFF
ECHO [+] START...

ECHO [+] Remove COLLECTOR OOOP-ml-templates-%1.oxt...
DEL OOOP-ml-templates-all.oxt

ECHO [+] Creating empty COLLECTOR OOOP-ml-templates-all.oxt...
COPY MultiLanguageTemplatePackage.oxt OOOP-ml-templates-all.oxt 

FOR %%L IN (hu de en-US fr it) DO CALL cretemex-ml.bat %%L
