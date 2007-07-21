#!/bin/bash
version=2.3.0.2
echo [+] START...
mkdir -p ../output

for mode in separated unified
do

for folder in templates templates-nonfree
do

  rm -fr ./_TEMP_ML_/
  mkdir -p ./_TEMP_ML_/

  for languages in hu de en-US it fr
  do
    
    echo ---------------------------------
    echo [!] Start building language: $languages
    echo ---------------------------------
    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

    for templates in educate finance forms layout misc officorr offimisc personal presnt
    do

      echo [+] Processing $templates folder, $languages language...
      if [ $mode == "unified" ]; then
      echo [.] UNIFIED mode
      mkdir -p ./_TEMP_/template/$templates
      crsync -avz --exclude-from=exclude.exc ../../extras/source/premium/$folder/$templates/lang/$languages/* ./_TEMP_/template/$templates/
      else
      echo [.] SEPARTATED  mode
      mkdir -p ./_TEMP_/template/$languages/$templates/
      rsync -avz --exclude-from=exclude.exc ../../extras/source/premium/$folder/$templates/lang/$languages/* ./_TEMP_/template/$languages/$templates/
      fi
      #  /EXCLUDE:cretwork.exc

    done
    rm -fr ./_TEMP_/*/.svn/*
    rm -fr ./_TEMP_/*makefile.mk
    rm -fr ./_TEMP_/*delzip
    rm -fr ./_TEMP_/*dummy.txt

    cp -fr ./_TEMP_/* ./_TEMP_ML_/
  
    echo [+] Adding licensing information...
    mkdir -p ./_TEMP_/licenses
    cp -f ../../documents/license/extension/license_*.txt ./_TEMP_/licenses/

    echo [+] Adding readme information...
    mkdir -p ./_TEMP_/readmes
    cp -f ../../documents/license/extension/readme_*.txt ./_TEMP_/readmes/

    echo [+] Adding description information...
    echo "   <identifier value=\""net.sf.ooop.oxygenoffice.$folder.$languages.$mode"\" />" > description.xml.middle
    echo "   <version value=\""$version"\" />" >> description.xml.middle

    cat description.xml.begin > ./_TEMP_/description.xml & cat description.xml.middle >> ./_TEMP_/description.xml & cat description.xml.end >> ./_TEMP_/description.xml
    rm -f description.xml.middle

    echo [+] Remove OOOP-$folder-$mode-$languages-$version.oxt
    rm -f ../output/OOOP-$folder-$mode-$languages-$version.oxt

    echo [+] Creating empty OOOP-$folder-$mode-$languages-$version.oxt...
    if [ $mode == "unified" ];
    then
    cp -f TemplatePackage.oxt ../output/OOOP-$folder-$mode-$languages-$version.oxt
    else
    cp -f MultiLanguageTemplatePackage.oxt ../output/OOOP-$folder-$mode-$languages-$version.oxt
    fi

    echo [+] Compressing to OOOP-$folder-$mode-$languages-$version.oxt...
    cd ./_TEMP_
    zip -r -9 ../../output/OOOP-$folder-$mode-$languages-$version.oxt *
    cd ..

    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

    echo ===============================================

  done

  echo [+] Adding licensing information...
  mkdir -p ./_TEMP_ML_/licenses
  cp -f ../../documents/license/extension/license_*.txt ./_TEMP_ML_/licenses/

  echo [+] Adding readme information...
  mkdir -p ./_TEMP_ML_/readmes
  cp -f ../../documents/license/extension/readme_*.txt ./_TEMP_ML_/readmes/

  echo [+] Adding description information...
  echo "   <identifier value=\""net.sf.ooop.oxygenoffice.gallery.free.multi_language.$mode"\" />" > description.xml.middle
  echo "   <version value=\""$version"\" />" >> description.xml.middle

  cat description.xml.begin > ./_TEMP_ML_/description.xml & cat description.xml.middle >> ./_TEMP_ML_/description.xml & cat description.xml.end >> ./_TEMP_ML_/description.xml
  rm -f description.xml.middle

  echo [+] Remove OOOP-$folder-$mode-all-$version.oxt...
  rm -f ../output/OOOP-$folder-$mode-all-$version.oxt

  echo [+] Creating empty OOOP-$folder-$mode-all-$version.oxt...
    if [ $mode == "unified" ];
    then
    cp -f TemplatePackage.oxt ../output/OOOP-$folder-$mode-all-$version.oxt
    else
    cp -f MultiLanguageTemplatePackage.oxt ../output/OOOP-$folder-$mode-all-$version.oxt
    fi

  cd ./_TEMP_ML_
  zip -r -9 ../../output/OOOP-$folder-$mode-all-$version.oxt *
  cd ..

  rm -fr ./_TEMP_ML_

done

done

echo [+] Finished, exiting and go home...

exit 0