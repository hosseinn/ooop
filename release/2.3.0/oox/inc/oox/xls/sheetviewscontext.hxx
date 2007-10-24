/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sheetviewscontext.hxx,v $
 *
 *  $Revision: 1.1.2.8 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:01 $
 *
 *  The Contents of this file are made available subject to
 *  the terms of GNU Lesser General Public License Version 2.1.
 *
 *
 *    GNU Lesser General Public License Version 2.1
 *    =============================================
 *    Copyright 2005 by Sun Microsystems, Inc.
 *    901 San Antonio Road, Palo Alto, CA 94303, USA
 *
 *    This library is free software; you can redistribute it and/or
 *    modify it under the terms of the GNU Lesser General Public
 *    License version 2.1, as published by the Free Software Foundation.
 *
 *    This library is distributed in the hope that it will be useful,
 *    but WITHOUT ANY WARRANTY; without even the implied warranty of
 *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *    Lesser General Public License for more details.
 *
 *    You should have received a copy of the GNU Lesser General Public
 *    License along with this library; if not, write to the Free Software
 *    Foundation, Inc., 59 Temple Place, Suite 330, Boston,
 *    MA  02111-1307  USA
 *
 ************************************************************************/

#ifndef OOX_XLS_SHEETVIEWSCONTEXT_HXX
#define OOX_XLS_SHEETVIEWSCONTEXT_HXX

#include "oox/xls/excelcontextbase.hxx"
#include "oox/xls/worksheethelper.hxx"

namespace oox {
namespace xls {

// ============================================================================

class OoxSheetViewsContext : public WorksheetContextBase
{
public:
    explicit            OoxSheetViewsContext( const WorksheetFragmentBase& rFragment );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );

private:
    void                importSheetView( const ::oox::core::AttributeList& rAttribs );
    void                importPane( const ::oox::core::AttributeList& rAttribs );
    void                importSelection( const ::oox::core::AttributeList& rAttribs );

private:
    OoxSheetViewData*   mpViewData;
};

// ============================================================================

// record identifiers ---------------------------------------------------------

const sal_uInt16 BIFF_ID_PANE               = 0x0041;
const sal_uInt16 BIFF_ID_SCL                = 0x00A0;
const sal_uInt16 BIFF_ID_SELECTION          = 0x001D;
const sal_uInt16 BIFF2_ID_WINDOW2           = 0x003E;
const sal_uInt16 BIFF3_ID_WINDOW2           = 0x023E;

// record constants -----------------------------------------------------------

const sal_uInt8 BIFF_PANE_BOTTOMRIGHT       = 0;        /// Bottom-right pane.
const sal_uInt8 BIFF_PANE_TOPRIGHT          = 1;        /// Right, or top-right pane.
const sal_uInt8 BIFF_PANE_BOTTOMLEFT        = 2;        /// Bottom, or bottom-left pane.
const sal_uInt8 BIFF_PANE_TOPLEFT           = 3;        /// Single, top, left, or top-left pane.

const sal_uInt16 BIFF_WIN2_SHOWFORMULAS     = 0x0001;
const sal_uInt16 BIFF_WIN2_SHOWGRID         = 0x0002;
const sal_uInt16 BIFF_WIN2_SHOWHEADINGS     = 0x0004;
const sal_uInt16 BIFF_WIN2_FROZEN           = 0x0008;
const sal_uInt16 BIFF_WIN2_SHOWZEROS        = 0x0010;
const sal_uInt16 BIFF_WIN2_DEFGRIDCOLOR     = 0x0020;
const sal_uInt16 BIFF_WIN2_RIGHTTOLEFT      = 0x0040;
const sal_uInt16 BIFF_WIN2_SHOWOUTLINE      = 0x0080;
const sal_uInt16 BIFF_WIN2_FROZENNOSPLIT    = 0x0100;
const sal_uInt16 BIFF_WIN2_SELECTED         = 0x0200;
const sal_uInt16 BIFF_WIN2_DISPLAYED        = 0x0400;
const sal_uInt16 BIFF_WIN2_PAGEBREAKMODE    = 0x0800;

// ----------------------------------------------------------------------------

class BiffSheetViewsContext : public WorksheetHelper
{
public:
    explicit            BiffSheetViewsContext( const WorksheetHelper& rSheetHelper );

    /** Tries to import a sheet view data record. */
    void                importRecord( BiffInputStream& rStrm );

private:
    void                importPane( BiffInputStream& rStrm );
    void                importScl( BiffInputStream& rStrm );
    void                importSelection( BiffInputStream& rStrm );
    void                importWindow2Biff2( BiffInputStream& rStrm );
    void                importWindow2Biff3( BiffInputStream& rStrm );

private:
    OoxSheetViewData&   mrViewData;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

