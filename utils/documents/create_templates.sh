#!/bin/bash
version=2.3.0.1
echo [+] START...
mkdir -p ../output

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
      mkdir -p ./_TEMP_/template/$templates
      cp -f ../../extras/source/premium/$folder/$templates/lang/$languages/*.* ./_TEMP_/template/$templates/
      #  /EXCLUDE:cretwork.exc

    done

    cp -fr ./_TEMP_/* ./_TEMP_ML_/
  
    echo [+] Adding licensing information...
    mkdir -p ./_TEMP_/licenses
    cp -f ../../documents/license/extension/license_*.txt ./_TEMP_/licenses/

    echo [+] Adding readme information...
    mkdir -p ./_TEMP_/readmes
    cp -f ../../documents/license/extension/readme_*.txt ./_TEMP_/readmes/

    echo [+] Adding description information...
    echo "   <identifier value=\""net.sf.ooop.oxygenoffice.gallery.free.$languages"\" />" > description.xml.middle
    echo "   <version value=\""$version"\" />" >> description.xml.middle

    cat description.xml.begin > ./_TEMP_/description.xml & cat description.xml.middle >> ./_TEMP_/description.xml & cat description.xml.end >> ./_TEMP_/description.xml
    rm -f description.xml.middle

    echo [+] Remove OOOP-$folder-$version.oxt...
    rm -f ../output/OOOP-$folder-$languages-$version.oxt

    echo [+] Creating empty OOOP-$folder-$languages-$version.oxt...
    cp -f TemplatePackage.oxt ../output/OOOP-$folder-$languages-$version.oxt

    echo [+] Compressing to OOOP-$folder-$languages-$version.oxt...
    cd ./_TEMP_
    zip -r -9 ../../output/OOOP-$folder-$languages-$version.oxt *
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
  echo "   <identifier value=\""net.sf.ooop.oxygenoffice.gallery.free.multi_language"\" />" > description.xml.middle
  echo "   <version value=\""$version"\" />" >> description.xml.middle

  cat description.xml.begin > ./_TEMP_ML_/description.xml & cat description.xml.middle >> ./_TEMP_ML_/description.xml & cat description.xml.end >> ./_TEMP_ML_/description.xml
  rm -f description.xml.middle

  echo [+] Remove OOOP-$folder-all-$version.oxt...
  rm -f ../output/OOOP-$folder-all-$version.oxt

  echo [+] Creating empty OOOP-$folder-all-$version.oxt...
  cp -f TemplatePackage.oxt ../output/OOOP-$folder-all-$version.oxt

  cd ./_TEMP_ML_
  zip -r -9 ../../output/OOOP-$folder-all-$version.oxt *
  cd ..

  rm -fr ./_TEMP_ML_

done


echo [+] Finished, exiting and go home...

exit 0