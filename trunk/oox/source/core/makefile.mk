#*************************************************************************
#
#   OpenOffice.org - a multi-platform office productivity suite
#
#   $RCSfile: makefile.mk,v $
#
#   $Revision: 1.1.2.10 $
#
#   last change: $Author: dr $ $Date: 2007/08/14 13:35:31 $
#
#   The Contents of this file are made available subject to
#   the terms of GNU Lesser General Public License Version 2.1.
#
#
#     GNU Lesser General Public License Version 2.1
#     =============================================
#     Copyright 2005 by Sun Microsystems, Inc.
#     901 San Antonio Road, Palo Alto, CA 94303, USA
#
#     This library is free software; you can redistribute it and/or
#     modify it under the terms of the GNU Lesser General Public
#     License version 2.1, as published by the Free Software Foundation.
#
#     This library is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#     Lesser General Public License for more details.
#
#     You should have received a copy of the GNU Lesser General Public
#     License along with this library; if not, write to the Free Software
#     Foundation, Inc., 59 Temple Place, Suite 330, Boston,
#     MA  02111-1307  USA
#
#*************************************************************************

PRJ=..$/..

PRJNAME=oox
TARGET=core
AUTOSEG=true

ENABLE_EXCEPTIONS=TRUE

# --- Settings -----------------------------------------------------

.INCLUDE :  settings.mk
.INCLUDE: $(PRJ)$/util$/makefile.pmk

# --- Files --------------------------------------------------------

SLOFILES =	\
		$(SLO)$/attributelist.obj			\
		$(SLO)$/binarycodec.obj				\
		$(SLO)$/binaryfilterbase.obj		\
		$(SLO)$/binaryinputstream.obj		\
		$(SLO)$/binaryoutputstream.obj		\
		$(SLO)$/containerhelper.obj			\
		$(SLO)$/context.obj					\
		$(SLO)$/facreg.obj					\
		$(SLO)$/filterbase.obj				\
		$(SLO)$/filterdetect.obj			\
		$(SLO)$/fragmenthandler.obj			\
		$(SLO)$/olestorage.obj				\
		$(SLO)$/propertymap.obj				\
		$(SLO)$/propertysequence.obj		\
		$(SLO)$/propertyset.obj				\
		$(SLO)$/relations.obj				\
		$(SLO)$/relationshandler.obj		\
		$(SLO)$/storagebase.obj				\
		$(SLO)$/xmlfilterbase.obj			\
		$(SLO)$/zipstorage.obj

# --- Targets -------------------------------------------------------

.INCLUDE :  target.mk