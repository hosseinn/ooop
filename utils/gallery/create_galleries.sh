#!/bin/bash
version=2.3.0.2
echo [+] START...
mkdir -p ../output

for folder in accessories accessories-nonfree
do

  rm -fr ./_TEMP_ML_/

    echo ---------------------------------
    echo [!] Start building galleries...
    echo ---------------------------------
    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

      echo [+] Processing $folder galleries...
      mkdir -p ./_TEMP_/gallery/
      rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../extras/source/gallery/$folder/* ./_TEMP_/gallery/

    echo [+] Adding licensing information...
    mkdir -p ./_TEMP_/licenses
    cp -f ../../documents/license/extension/license_*.txt ./_TEMP_/licenses/

    echo [+] Adding readme information...
    mkdir -p ./_TEMP_/readmes
    cp -f ../../documents/license/extension/readme_*.txt ./_TEMP_/readmes/

    echo [+] Adding description information...
    echo "   <identifier value=\""net.sf.ooop.oxygenoffice.$folder"\" />" > description.xml.middle
    echo "   <version value=\""$version"\" />" >> description.xml.middle

    cat ./description.xml.begin > ./_TEMP_/description.xml
    cat ./description.xml.middle >> ./_TEMP_/description.xml
    cat ./description.xml.end >> ./_TEMP_/description.xml
    rm -f description.xml.middle

    echo [+] Remove OOOP-$folder-$version.oxt
    rm -f ../output/OOOP-$folder-$version.oxt

    echo [+] Creating empty OOOP-$folder-$version.oxt...
    cp -f GalleryPackage.oxt ../output/OOOP-$folder-$version.oxt

    echo [+] Compressing to OOOP-$folder-$version.oxt...
    cd ./_TEMP_
    zip -r -9 ../../output/OOOP-$folder-$version.oxt *
    cd ..

    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

    echo ===============================================

done


echo [+] Finished, exiting and go home...

exit 0