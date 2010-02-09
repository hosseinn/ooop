#!/bin/bash
version=20100206
echo [+] START...
mkdir -p ../output

rm -fr ./_TEMP_
    mkdir -p ./_TEMP_/dicts/
    rsync -avv --progress --human-readable --exclude-from=exclude.exc ../../dictionaries/* ./_TEMP_/dictionaries/
    echo [+] Compressing to extensionaids-$version.zip...
    cd ./_TEMP_/dictionaries
    zip -r -9 ../../../output/extensionaids-$version.zip *
    cd ../../


    echo [+] Remove TEMP...
    rm -fr ./_TEMP_

echo [+] Finished, exiting and go home...

exit 0
