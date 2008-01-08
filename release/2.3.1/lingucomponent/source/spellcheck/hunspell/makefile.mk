#*************************************************************************
#
#   $RCSfile: makefile.mk,v $
#
#   $Revision: 1.3.52.7 $
#
#   last change: $Author: nemeth $ $Date: 2007/09/25 08:12:06 $
#
#   The Contents of this file are made available subject to
#   the terms of GNU Lesser General Public License Version 2.1.
#
#   GNU Lesser General Public License Version 2.1
#   =============================================
#   Copyright 2005 by Sun Microsystems, Inc.
#   901 San Antonio Road, Palo Alto, CA 94303, USA
#
#   This library is free software; you can redistribute it and/or
#   modify it under the terms of the GNU Lesser General Public
#   License version 2.1, as published by the Free Software Foundation.
#
#   This library is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#   Lesser General Public License for more details.
#
#   You should have received a copy of the GNU Lesser General Public
#   License along with this library; if not, write to the Free Software
#   Foundation, Inc., 59 Temple Place, Suite 330, Boston,
#   MA  02111-1307  USA
#
#*************************************************************************

PRJ = ..$/..$/..

PRJNAME	= lingucomponent
TARGET	= hunspell
USE_DEFFILE = TRUE



#----- Settings ---------------------------------------------------------

.INCLUDE : settings.mk

# --- Files --------------------------------------------------------

.IF "$(SYSTEM_HUNSPELL)" == "YES"
@all:
	@echo "Using system hunspell..."
.ENDIF

# all_target: ALLTAR DICTIONARY
all_target: ALLTAR

.IF "$(GUI)" == "UNX"
CDEFS+=-DOPENOFFICEORG
.ENDIF

SLOFILES=	\
		$(SLO)$/affentry.obj \
		$(SLO)$/affixmgr.obj \
		$(SLO)$/csutil.obj \
		$(SLO)$/phonet.obj \
		$(SLO)$/hashmgr.obj \
		$(SLO)$/suggestmgr.obj \
		$(SLO)$/hunspell.obj

SHL1TARGET= 	$(TARGET)

.IF "$(GUI)" == "UNX"
SHL1STDLIBS= 	$(ICUUCLIB)
.ELSE
SHL1STDLIBS=
.ENDIF

# build DLL
SHL1DEPN=
SHL1IMPLIB=	i$(TARGET)
SHL1LIBS=	$(SLB)$/$(TARGET).lib
SHL1DEF=	$(MISC)$/$(SHL1TARGET).def
DEF1NAME=	$(SHL1TARGET)
.IF "$(GUI)$(COM)"=="WNTGCC"
DEFLIB1NAME =$(TARGET)
.ENDIF


# --- Targets ------------------------------------------------------

.INCLUDE : target.mk

