#!/bin/bash
# root of build tree
my_BUILD_ROOT="/cygdrive/c/ooo2"


myIDE_PATH="/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/IDE"
myMSPDB_PATH="/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/IDE"
myFWSDK_PATH="/cygdrive/c/Program Files/Microsoft.NET/SDK/v1.1"
myJSDK_PATH="/cygdrive/c/j2sdk1.4.2_10"
myPSDK_PATH="/cygdrive/c/Program Files/Microsoft Platform SDK"
myDX_PATH="/cygdrive/c/DXSDK"
myANTHOME="/cygdrive/c/apache-ant-1.6.5"
myCLHOME="/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Vc7"
myNMAKE_PATH="/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Vc7/bin"
myCSCPATH="/cygdrive/c/WINDOWS/Microsoft.NET/Framework/v1.1.4322"
myMIDLPATH="/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/Tools/Bin"
myNSISPATH="/cygdrive/c/Program Files/NSIS"

PATH "${PATH}:${myIDE_PATH}"

# desired languages
my_LANGUAGES="hu de fr it tr ka fi pl nl pt-BR es ja zh-CN"

BUILDNUMBER="OxygenOffice Professional 2.3.0 - OxygenOffice Build 2"

BUILDER="OxygenOffice Professional Team"

./configure --with-lang="$my_LANGUAGES" --with-mspdb-path="$myMSPDB_PATH" --with-frame-home="$myFWSDK_PATH" --with-jdk-home="$myJSDK_PATH" --with-use-shell=bash --with-psdk-home="$myPSDK_PATH" --with-directx-home="$myDX_PATH" --with-ant-home=$myANTHOME --with-cl-home="$myCLHOME" --with-nmake-path="$myNMAKE_PATH" --with-csc-path="$myCSCPATH" --with-nsis-path="$myNSISPATH"  --with-build-version="$BUILDNUMBER" --with-vendor="$BUILDER" --disable-build-mozilla --enable-vba --with-vba-package-format="builtin"

# supply temp dir
echo setenv TEMP /tmp >>../winenv.set
echo setenv TMP /tmp >>../winenv.set

