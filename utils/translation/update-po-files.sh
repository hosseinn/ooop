#!/bin/bash

# Update supported languages from properties file to po and pot files

CURRENT_DIR=`pwd`

PROPERTIES_DIR="$CURRENT_DIR/dialogs"
PO_DIR="$CURRENT_DIR/po"
TEMP_DIR="$CURRENT_DIR/temp"

MSGCAT_OPTION="--use-first"

mkdir -p $PO_DIR/template $TEMP_DIR/template

for language in `cat $PO_DIR/supported_languages`
do
	cp $PROPERTIES_DIR/*_$language.properties $TEMP_DIR/
	cp $PROPERTIES_DIR/*en_US.properties $TEMP_DIR/template
	cd $TEMP_DIR/template
	for i in *_en_US.properties
	do
		j=`echo $i | sed "s/en_US/$language/g"`
		mv "$i" "$j"
	done
done

moz2po --input $PROPERTIES_DIR --output $TEMP_DIR --template $TEMP_DIR/template --exclude .default

# Checking for supported languages
for language in `cat $PO_DIR/supported_languages`
do
# Make language dependent directories and move a po files there
	echo "[+] Now processing $language language..."
	mkdir -p $PO_DIR/$language
	cd $TEMP_DIR
	for po_file in *_$language.properties.po
	do
    		if test -f "$PO_DIR/$language/$po_file"
		then
        		echo "[.] Merging ($language) $po_file..."
        		po_temp=`mktemp /tmp/$po_file.XXXXXX`
            		msgcat $MSGCAT_OPTION "$PO_DIR/$language/$po_file" "$po_file" > $po_temp
		        if test "$?" = "0"
			then
            			mv $po_temp "$PO_DIR/$language/$po_file"
            			chmod 644 "$PO_DIR/$language/$po_file"
            			echo "[.] $po_file was succesfully updated"
        		else
            			echo "[!] $po_file was not updated"
        		fi
    		else
        		echo "[+] Adding $po_file..."
        		cp "$po_file" "$PO_DIR/$language/$po_file"
    		fi
	done
# Make language dependent templates and move a pot files to template folder
done

cd $CURRENT_DIR

moz2po -P -i $PROPERTIES_DIR/ -o $TEMP_DIR/
mv $TEMP_DIR/*_en_US.properties.pot $PO_DIR/template

rm -fr $TEMP_DIR
