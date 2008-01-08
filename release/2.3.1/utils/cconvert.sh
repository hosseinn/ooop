#!/bin/bash
export lcl="tr"
export lco="TR"
for file in *.ot?
do
mkdir -p ./work/"$file"
mkdir -p ./output/$lcl
unzip -o "$file" -d ./work/"$file"
cd ./work/"$file"
echo 1
sed "s/<dc:language>[a-z][a-z]-[A-Z][A-Z]<\/dc:language>/<dc:language>$lcl-$lco<\/dc:language>/g" meta.xml  > new_meta.xml
echo 2
sed "s/<config:config-item config:name=\"Language\" config:type=\"string\">[a-z][a-z]<\/config:config-item>/<config:config-item config:name=\"Language\" config:type=\"string\">$lcl<\/config:config-item>/g"  settings.xml  > settings_1.xml
echo 3
sed "s/<config:config-item config:name=\"Country\" config:type=\"string\">[A-Z][A-Z]<\/config:config-item>/<config:config-item config:name=\"Country\" config:type=\"string\">$lco<\/config:config-item>/g"  settings_1.xml  > new_settings.xml
echo 4
sed "s/fo:language=\"[a-z][a-z]\"/fo:language=\"$lcl\"/g" styles.xml  > styles_1.xml
echo 5
sed "s/fo:country=\"[A-Z][A-Z]\"/fo:country=\"$lco\"/g" styles_1.xml  > styles.xml
echo 6
sed "s/style:language-asian=\"[a-z][a-z]\"/style:language-asian=\"none\"/g" styles.xml  > styles_1.xml
echo 7
sed "s/style:country-asian=\"[A-Z][A-Z]\"/style:country-asian=\"none\"/g" styles_1.xml  > styles.xml
echo 8
sed "s/style:language-complex=\"[a-z][a-z]\"/style:language-complex=\"none\"/g" styles.xml  > styles_1.xml
echo 9
sed "s/style:country-complex=\"[A-Z][A-Z]\"/style:country-complex=\"none\"/g" styles_1.xml  > new_styles.xml

mv new_meta.xml meta.xml
rm settings_1.xml
rm styles_1.xml
mv new_settings.xml settings.xml
mv new_styles.xml styles.xml
zip -r -9 ../../output/$lcl/"$file" *
cd ../..
rm -fr ./work/"$file"
done
