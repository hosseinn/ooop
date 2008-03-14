#!/bin/bash
#
for file in *.jpg
do
echo [+] Creating layout from $file...
mkdir -p ./work/$file
echo [+] Unpack master layout template...
unzip -o PRESENTATIONSAMPLE.otp -d ./work/$file/
echo [+] Converting $file to background image...
myrandom=$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM$RANDOM
picturename=`echo $myrandom | cut -b 11-42`
echo [.] Generated filename is: [$picturename]
convert $file -resize 1024x768 ./work/$file/Pictures/$picturename.png
echo [+] Converting $file to thumbnail image...
#256x192
convert $file -resize 256x192 ./work/$file/Thumbnails/thumbnail.png
echo [+] Starting text processing...
cd ./work/$file
echo [.] Change title...
sed "s/<dc:title>\*\*PRESENTATIONTITLE\*\*<\/dc:title>/<dc:title>$file<\/dc:title>/g" meta.xml  > new_meta.xml
mv new_meta.xml meta.xml
echo [.] Change content...
sed "s/PRESENTATIONSAMPLE/$file/g" content.xml  > new_content.xml
mv new_content.xml content.xml
echo [.] Change styles...
sed "s/PRESENTATIONSAMPLE/$file/g" styles.xml  > new_styles.xml
sed "s/1000000000000640000004B02C5236E4.png/$picturename.png/g" new_styles.xml > styles.xml
#sed "s/10000000000000400000004077CDC8F9/$picturename/g" styles.xml  > new_styles.xml
rm new_styles.xml
echo [.] Change manifest...
sed "s/1000000000000640000004B02C5236E4.png/$picturename.png/g" ./META-INF/manifest.xml > ./META-INF/new_manifest.xml
mv ./META-INF/new_manifest.xml ./META-INF/manifest.xml
echo [.] Remove old image...
rm ./Pictures/1000000000000640000004B02C5236E4.*
echo [+] Creating $file template...
zip -mr9 ../$file.otp * 
cd ../..
rm -fr ./work/$file/
done
