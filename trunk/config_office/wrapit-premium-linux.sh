#!/bin/bash

# We will use ccache on GNU/Linux
export CC="/usr/bin/ccache /usr/bin/gcc"
export CXX="/usr/bin/ccache /usr/bin/g++"

# Package formats
PKGFMTS="rpm deb"

# Ant home
ANTHOME="/usr/local/apache-ant-1.6.2"

# Java home
JAVAHOME="/usr/local/j2sdk1.4.2_10"

# Build number
BUILDNUMBER="OxygenOffice Professional 2.0.4 Premium FSF.hu Build 1"

# desired languages
MYLANGUAGES="hu de fr it tr ka"

# Configure flags
CONFIGUREFLAGS="--enable-kde --enable-gtk --enable-build-mozilla"

./configure --with-lang="${MYLANGUAGES}" --with-jdk-home=${JAVAHOME} --with-ant-home=${ANTHOME} --with-package-format="${PKGFMTS}" ${CONFIGUREFLAGS} --with-build-version="${BUILDNUMBER}"
