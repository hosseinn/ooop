#!/bin/bash
version=2.6.0.2
echo [+] START...
mkdir -p ../output

rm -fr ./_TEMP_
for folder in templates templates-nonfree
do
      mkdir -p ./_TEMP_/templates/premium/$folder/
      rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../extras/source/premium/$folder/* ./_TEMP_/templates/premium/$folder/
done

for folder in samples samples-nonfree
do
      mkdir -p ./_TEMP_/samples/premium/$folder/
      rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../extras/source/premium/$folder/* ./_TEMP_/samples/premium/$folder/
done

for folder in gallery
do
      mkdir -p ./_TEMP_/gallery/$folder/
      rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../extras/source/$folder/* ./_TEMP_/gallery/$folder/
done

for folder in truetype
do
      mkdir -p ./_TEMP_/fonts/$folder/
      rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../extras/source/$folder/* ./_TEMP_/fonts/$folder/
done

for package in templates samples gallery fonts
do
    echo [+] Compressing to OOOP-$package-pack-$version.zip...
    cd ./_TEMP_/$package
    zip -r -9 ../../../output/OOOP-$package-pack-$version.zip *
    cd ../../

done

    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

echo [+] Finished, exiting and go home...

exit 0
