/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: workbookfragment.hxx,v $
 *
 *  $Revision: 1.1.2.25 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/03 13:47:43 $
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

#ifndef OOX_XLS_WORKBOOKFRAGMENT_HXX
#define OOX_XLS_WORKBOOKFRAGMENT_HXX

#include <rtl/ustrbuf.hxx>
#include "oox/xls/excelfragmentbase.hxx"
#include "oox/xls/defnamesbuffer.hxx"

namespace oox {
namespace xls {

// ============================================================================

class OoxWorkbookFragment : public GlobalFragmentBase
{
public:
    explicit            OoxWorkbookFragment(
                            const GlobalDataHelper& rGlobalData,
                            const ::rtl::OUString& rFragmentPath );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    void                importWorkbookView( const ::oox::core::AttributeList& rAttribs );
    void                importExternalReference( const ::oox::core::AttributeList& rAttribs );
    void                importPivotCache( const ::oox::core::AttributeList& rAttribs );

    void                finalizeImport();

private:
    DefinedNameRef      mxCurrName;
};

// ============================================================================

// record identifiers ---------------------------------------------------------

const sal_uInt16 BIFF_ID_BOUNDSHEET         = 0x0085;
const sal_uInt16 BIFF_ID_CODEPAGE           = 0x0042;
const sal_uInt16 BIFF_ID_CRN                = 0x005A;
const sal_uInt16 BIFF2_ID_EXTERNNAME        = 0x0023;
const sal_uInt16 BIFF3_ID_EXTERNNAME        = 0x0223;
const sal_uInt16 BIFF5_ID_EXTERNNAME        = 0x0023;
const sal_uInt16 BIFF_ID_EXTERNSHEET        = 0x0017;
const sal_uInt16 BIFF_ID_EXTSST             = 0x00FF;
const sal_uInt16 BIFF_ID_FILEPASS           = 0x002F;
const sal_uInt16 BIFF2_ID_FONT              = 0x0031;
const sal_uInt16 BIFF3_ID_FONT              = 0x0231;
const sal_uInt16 BIFF5_ID_FONT              = 0x0031;
const sal_uInt16 BIFF_ID_FONTCOLOR          = 0x0045;
const sal_uInt16 BIFF2_ID_FORMAT            = 0x001E;
const sal_uInt16 BIFF4_ID_FORMAT            = 0x041E;
const sal_uInt16 BIFF2_ID_NAME              = 0x0018;
const sal_uInt16 BIFF3_ID_NAME              = 0x0218;
const sal_uInt16 BIFF5_ID_NAME              = 0x0018;
const sal_uInt16 BIFF_ID_PALETTE            = 0x0092;
const sal_uInt16 BIFF_ID_PROJEXTSHEET       = 0x00A3;
const sal_uInt16 BIFF_ID_SHEETHEADER        = 0x008F;
const sal_uInt16 BIFF_ID_SST                = 0x00FC;
const sal_uInt16 BIFF_ID_STYLE              = 0x0293;
const sal_uInt16 BIFF_ID_SUPBOOK            = 0x01AE;
const sal_uInt16 BIFF_ID_WINDOW1            = 0x003D;
const sal_uInt16 BIFF_ID_XCT                = 0x0059;
const sal_uInt16 BIFF2_ID_XF                = 0x0043;
const sal_uInt16 BIFF3_ID_XF                = 0x0243;
const sal_uInt16 BIFF4_ID_XF                = 0x0443;
const sal_uInt16 BIFF5_ID_XF                = 0x00E0;

// record constants -----------------------------------------------------------

const sal_uInt16 BIFF_FILEPASS_BIFF2        = 0x0000;
const sal_uInt16 BIFF_FILEPASS_BIFF8        = 0x0001;
const sal_uInt16 BIFF_FILEPASS_BIFF8_RCF    = 0x0001;
const sal_uInt16 BIFF_FILEPASS_BIFF8_STRONG = 0x0002;

const sal_uInt16 BIFF_WIN1_HIDDEN           = 0x0001;
const sal_uInt16 BIFF_WIN1_MINIMIZED        = 0x0002;
const sal_uInt16 BIFF_WIN1_SHOWHORSCROLL    = 0x0008;
const sal_uInt16 BIFF_WIN1_SHOWVERSCROLL    = 0x0010;
const sal_uInt16 BIFF_WIN1_SHOWTABBAR       = 0x0020;

// ----------------------------------------------------------------------------

class BiffWorkbookFragment : public GlobalDataHelper
{
public:
    explicit            BiffWorkbookFragment( const GlobalDataHelper& rGlobalData );

    /** Imports the entire workbook stream, including all contained worksheets. */
    bool                importWorkbook( BiffInputStream& rStrm );

private:
    /** Imports the workbook globals fragment from current stream position. */
    bool                importGlobalsFragment( BiffInputStream& rStrm );
    /** Imports a complete BIFF4 workspace fragment (with embedded sheets). */
    bool                importWorkspaceFragment( BiffInputStream& rStrm );

    /** Imports a sheet fragment with passed type from current stream position. */
    bool                importSheetFragment(
                            BiffInputStream& rStrm,
                            BiffFragmentType eFragment,
                            sal_Int32 nSheet );
    /** Imports a worksheet fragment from current stream position. */
    bool                importWorksheetFragment(
                            BiffInputStream& rStrm,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );
    /** Imports a chartsheet fragment from current stream position. */
    bool                importChartsheetFragment(
                            BiffInputStream& rStrm,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );
    /** Imports a macrosheet fragment from current stream position. */
    bool                importMacrosheetFragment(
                            BiffInputStream& rStrm,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );

    /** Imports the FILEPASS record and sets a decoder at the passed stream. */
    bool                importFilePass( BiffInputStream& rStrm );
    /** Imports the WINDOW1 record containing global view settings. */
    void                importWindow1( BiffInputStream& rStrm );
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

