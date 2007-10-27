@ECHO OFF
IF %1=="" GOTO EXIT
FOR %%D IN (samples samples-nonfree)     DO MD %%D\lang\%1 & COPY .nametranslation.table %%D\lang\%1\.nametranslation.table
FOR %%D IN (samples samples-nonfree)     DO FOR %%S IN (advertisement documentation) DO MD %%D\%%S\lang\%1 & COPY dummy.txt %%D\%%S\lang\%1\dummy_%%D.txt
FOR %%D IN (templates templates-nonfree) DO MD %%D\lang\%1
FOR %%D IN (templates templates-nonfree) DO FOR %%S IN (educate finance forms layout misc officorr offimisc personal presnt) DO MD %%D\%%S\lang\%1 & COPY dummy.txt %%D\%%S\lang\%1\dummy_%%D.txt
:EXIT
ECHO Please specify language code like this:
ECHO %0 hu