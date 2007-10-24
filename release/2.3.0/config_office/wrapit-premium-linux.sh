#!/bin/bash

# We will use ccache on GNU/Linux
export CC="/usr/bin/ccache /usr/bin/gcc"
export CXX="/usr/bin/ccache /usr/bin/g++"

# Package formats
PKGFMTS="rpm deb"

# Ant home
ANTHOME="/usr/local/apache-ant-1.6.5"

# Java home
JAVAHOME="/usr/lib/jvm/java-1.5.0-sun-1.5.0.11/"

# Build number
BUILDNUMBER="OxygenOffice Professional 2.3.0 - OxygenOffice Build 2"

# Builder name
BUILDER="OxygenOffice Professional Team"

# desired languages
MYLANGUAGES="hu de fr it tr ka fi pl nl pt-BR es"

# Configure flags
CONFIGUREFLAGS="--enable-kde --enable-gtk --enable-build-mozilla --with-system-nas  --enable-vba --with-vba-package-format=builtin --enable-evolution2 --enable-lockdown --enable-cairo"

./configure --with-lang="${MYLANGUAGES}" --with-jdk-home=${JAVAHOME} --with-ant-home=${ANTHOME} --with-package-format="${PKGFMTS}" ${CONFIGUREFLAGS} --with-build-version="${BUILDNUMBER}" --with-vendor="${BUILDER}"
