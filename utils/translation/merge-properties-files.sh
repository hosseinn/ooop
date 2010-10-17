#!/bin/bash

# Update properties file from PO files

SCRIPTNAME="Update properties files from PO files"
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

mkdir -p $TEMP_DIR/properties
cd $TEMP_DIR/properties/

for language in $SUPPORTED_LANGUAGES
do
    echo "[.] Processing $language language."
    mkdir -p $TEMP_DIR/$language/
    cp $PROPERTIES_DIR/*_en_US.properties $TEMP_DIR/$language/
    cd $TEMP_DIR/$language/
    if [ "$language" != "en_US" ]
    then
	for i in *.properties
	do
	    j=`echo $i | sed "s/en_US/$language/g"`
	    mv "$i" "$j"
	done
    fi
    po2moz --input $PO_DIR/$language/ --output $TEMP_DIR/properties/ --template $TEMP_DIR/$language/ --exclude "*.xdl" --exclude "*.default"
    cd $TEMP_DIR/properties/
    for file in *.properties
    do
	echo "[+] Finalizing $file file."
	uni2ascii -a U  "$file" > "$PROPERTIES_DIR/$file"
    done
    rm -f $file
done

rm -fr $TEMP_DIR
