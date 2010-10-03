#!/bin/bash

# Update supported languages from properties file to po and pot files

SCRIPTNAME="Update PO files from properties files"
VERSION="V1.0.0"
SCRIPTFILE=`basename $0`

echo -e "\n$SCRIPTNAME - $VERSION\n"

# Help, I need somebody
function USAGE ()
{
    echo ""
    echo "USAGE: "
    echo "    $SCRIPTFILE [SWITCH...]"
    echo ""
    echo "OPTIONS:"
    echo "        -p, --properties-dir DIRECTORY - Location of properties dir, it will converted to po file"
    echo "        -po, --po-dir DIRECTORY        - Location of po dir"
    echo "        -t, --temp DIRECTORY           - Location of temporary folder"
    echo "        -?, -h, --help                 - Displays this help"
    echo ""
    echo "EXAMPLE:"
    echo "    $SCRIPTFILE"
    echo ""
    exit $E_OPTERROR
}

# Set default setting - customer profile

CURRENT_DIR=`pwd`
PROPERTIES_DIR="$CURRENT_DIR/dialogs"
PO_DIR="$CURRENT_DIR/po"
TEMP_DIR="$CURRENT_DIR/temp"

# Process arguments
while [ $# -gt 0 ]
do
    case "$1" in
        -p|--properties-dir) PROPERTIES_DIR="$2"; shift;;
        -po|--po-dir) PO_DIR="$2"; shift;;
        -t|--temp-dir) TEMP_DIR="$2"; shift;;
        -?|-h|--help) USAGE;;
	*) echo -e "\nERROR: Unknown parameter\n"; USAGE; exit $E_OPTERROR;;
    esac
    shift
done

echo -e "Current settings:\n-------------------------------------------------------------"
echo    "Properties directory      : $PROPERTIES_DIR"
echo    "PO directory              : $PO_DIR"
echo -e "TEMP directory            : $TEMP_DIR\n"

echo "[.] Starting conversion..."

SUPPORTED_LANGUAGES="`cat $PO_DIR/supported_languages` en_US"

MSGCAT_OPTION="--use-first"

mkdir -p $PO_DIR/template $TEMP_DIR/template

for language in $SUPPORTED_LANGUAGES
do
	echo "[+] Preparing $language language..."
	cp $PROPERTIES_DIR/*_$language.properties $TEMP_DIR/
	cp $PROPERTIES_DIR/*en_US.properties $TEMP_DIR/template
	if [ "$language" != "en_US" ]
	then
		cd $TEMP_DIR/template
		for i in *_en_US.properties
		do
			j=`echo $i | sed "s/en_US/$language/g"`
			mv "$i" "$j"
		done
	fi
done

moz2po --input $PROPERTIES_DIR --output $TEMP_DIR --template $TEMP_DIR/template --exclude "*.xdl" --exclude "*.default"

# Checking for supported languages
for language in $SUPPORTED_LANGUAGES
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
            			echo "[.] $po_file was succesfully updated."
        		else
            			echo "[!] $po_file was not updated."
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
