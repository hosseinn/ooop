#*************************************************************************
#
#   OpenOffice.org - a multi-platform office productivity suite
#
#   $RCSfile: makefile.mk,v $
#
#   $Revision: 1.1.2.35 $
#
#   last change: $Author: dr $ $Date: 2007/09/05 12:31:23 $
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
TARGET=xls
AUTOSEG=true

ENABLE_EXCEPTIONS=TRUE

# --- Settings -----------------------------------------------------

.INCLUDE :  settings.mk
.INCLUDE: $(PRJ)$/util$/makefile.pmk

# --- Files --------------------------------------------------------

SLOFILES =									\
		$(SLO)$/addressconverter.obj		\
		$(SLO)$/autofiltercontext.obj		\
		$(SLO)$/biffcodec.obj				\
		$(SLO)$/biffformatdetector.obj		\
		$(SLO)$/biffhelper.obj				\
		$(SLO)$/biffinputstream.obj			\
		$(SLO)$/biffoutputstream.obj		\
		$(SLO)$/condformatbuffer.obj		\
		$(SLO)$/condformatcontext.obj		\
		$(SLO)$/connectionsfragment.obj		\
		$(SLO)$/contexthelper.obj			\
		$(SLO)$/datavalidationscontext.obj	\
		$(SLO)$/defnamesbuffer.obj			\
		$(SLO)$/dxfscontext.obj				\
		$(SLO)$/excelcontextbase.obj		\
		$(SLO)$/excelfilter.obj				\
		$(SLO)$/excelfragmentbase.obj		\
		$(SLO)$/externallinkbuffer.obj		\
		$(SLO)$/externallinkfragment.obj	\
		$(SLO)$/formulabase.obj				\
		$(SLO)$/formulaparser.obj			\
		$(SLO)$/globaldatahelper.obj		\
		$(SLO)$/headerfooterparser.obj		\
		$(SLO)$/numberformatsbuffer.obj		\
		$(SLO)$/pagestyle.obj				\
		$(SLO)$/pivotcachefragment.obj		\
		$(SLO)$/pivottablebuffer.obj		\
		$(SLO)$/pivottablefragment.obj		\
		$(SLO)$/querytablefragment.obj		\
		$(SLO)$/richstring.obj				\
		$(SLO)$/richstringcontext.obj		\
		$(SLO)$/sharedstringsbuffer.obj		\
		$(SLO)$/sharedstringsfragment.obj	\
		$(SLO)$/sheetcellrangemap.obj		\
		$(SLO)$/sheetdatacontext.obj		\
		$(SLO)$/sheetviewscontext.obj		\
		$(SLO)$/stylesbuffer.obj			\
		$(SLO)$/stylesfragment.obj			\
		$(SLO)$/stylespropertyhelper.obj	\
		$(SLO)$/themebuffer.obj				\
		$(SLO)$/tokenmapper.obj				\
		$(SLO)$/unitconverter.obj			\
		$(SLO)$/viewsettings.obj			\
		$(SLO)$/webquerybuffer.obj			\
		$(SLO)$/workbookfragment.obj		\
		$(SLO)$/worksheetbuffer.obj			\
		$(SLO)$/worksheetfragment.obj		\
		$(SLO)$/worksheethelper.obj

# --- Targets -------------------------------------------------------

.INCLUDE :  target.mk
