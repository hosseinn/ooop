# root of build tree
setenv my_BUILD_ROOT /cygdrive/c/ooo2


setenv myIDE_PATH "/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/IDE"
setenv myMSPDB_PATH "/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/IDE"
setenv myFWSDK_PATH "/cygdrive/c/Program Files/Microsoft.NET/SDK/v1.1"
setenv myJSDK_PATH "/cygdrive/c/j2sdk1.4.2_10"
setenv myPSDK_PATH "/cygdrive/c/Program Files/Microsoft Platform SDK"
setenv myDX_PATH "/cygdrive/c/DXSDK"
setenv myANTHOME "/cygdrive/c/apache-ant-1.6.5"
setenv myCLHOME "/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Vc7"
setenv myNMAKE_PATH "/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Vc7/bin"
setenv myCSCPATH "/cygdrive/c/WINDOWS/Microsoft.NET/Framework/v1.1.4322"
setenv myMIDLPATH "/cygdrive/c/Program Files/Microsoft Visual Studio .NET 2003/Common7/Tools/Bin" 
setenv myNSISPATH "/cygdrive/c/Program Files/NSIS"

setenv PATH "${PATH}:${myIDE_PATH}"

# desired languages
setenv my_LANGUAGES "hu de fr it tr ka"

setenv BUILDNUMBER "OxygenOffice Professional 2.2 - OxygenOffice Build 1"

./configure --with-lang="$my_LANGUAGES" --with-mspdb-path="$myMSPDB_PATH" --with-frame-home="$myFWSDK_PATH" --with-jdk-home="$myJSDK_PATH" --with-use-shell=tcsh --with-psdk-home="$myPSDK_PATH" --with-directx-home="$myDX_PATH" --with-ant-home=$myANTHOME --with-cl-home="$myCLHOME" --with-nmake-path="$myNMAKE_PATH" --with-csc-path="$myCSCPATH" --with-nsis-path="$myNSISPATH"  --with-build-version="$BUILDNUMBER"

# supply temp dir
echo setenv TEMP /tmp >>../winenv.set
echo setenv TMP /tmp >>../winenv.set

