/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: worksheetfragment.hxx,v $
 *
 *  $Revision: 1.1.2.38 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:58:00 $
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

#ifndef OOX_XLS_WORKSHEETFRAGMENT_HXX
#define OOX_XLS_WORKSHEETFRAGMENT_HXX

#include "oox/xls/excelfragmentbase.hxx"
#include "oox/xls/worksheethelper.hxx"

namespace oox {
namespace xls {

// ============================================================================

class OoxWorksheetFragment : public WorksheetFragmentBase
{
public:
    explicit            OoxWorksheetFragment(
                            const GlobalDataHelper& rGlobalData,
                            const ::rtl::OUString& rFragmentPath,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        onCreateContext( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    /** Initial processing on the root worksheet element. */
    void                importWorksheet( const ::oox::core::AttributeList& rAttribs );
    /** Imports column settings from a col element. */
    void                importColumn( const ::oox::core::AttributeList& rAttribs );
    /** Imports a merged cell range from a mergeCell element. */
    void                importMergeCell( const ::oox::core::AttributeList& rAttribs );
    /** Imports sheet properties from a sheetPr element. */
    void                importSheetPr( const ::oox::core::AttributeList& rAttribs );
    /** Imports outline properties from an outlinePr element. */
    void                importOutlinePr( const ::oox::core::AttributeList& rAttribs );
    /** Imports page settings from a pageSetUpPr element. */
    void                importPageSetUpPr( const ::oox::core::AttributeList& rAttribs );
    /** Imports sheet format properties. */
    void                importSheetFormatPr( const ::oox::core::AttributeList& rAttribs );
    /** Imports individual break that is either within row or column break context. */
    void                importBrk( const ::oox::core::AttributeList& rAttribs );
    /** Imports the hyperlink element containing a hyperlink for a cell range. */
    void                importHyperlink( const ::oox::core::AttributeList& rAttribs );
};

// ============================================================================

// record identifiers ---------------------------------------------------------

const sal_uInt16 BIFF_ID_BOTTOMMARGIN       = 0x0029;
const sal_uInt16 BIFF_ID_COLINFO            = 0x007D;
const sal_uInt16 BIFF_ID_COLUMNDEFAULT      = 0x0020;
const sal_uInt16 BIFF_ID_COLWIDTH           = 0x0024;
const sal_uInt16 BIFF_ID_DEFCOLWIDTH        = 0x0055;
const sal_uInt16 BIFF2_ID_DEFROWHEIGHT      = 0x0025;
const sal_uInt16 BIFF3_ID_DEFROWHEIGHT      = 0x0225;
const sal_uInt16 BIFF_ID_FOOTER             = 0x0015;
const sal_uInt16 BIFF_ID_HCENTER            = 0x0083;
const sal_uInt16 BIFF_ID_HEADER             = 0x0014;
const sal_uInt16 BIFF_ID_HLINK              = 0x01B8;
const sal_uInt16 BIFF_ID_HORPAGEBREAKS      = 0x001B;
const sal_uInt16 BIFF_ID_LEFTMARGIN         = 0x0026;
const sal_uInt16 BIFF_ID_MERGEDCELLS        = 0x00E5;
const sal_uInt16 BIFF_ID_PRINTHEADERS       = 0x002A;
const sal_uInt16 BIFF_ID_PRINTGRIDLINES     = 0x002B;
const sal_uInt16 BIFF_ID_RIGHTMARGIN        = 0x0027;
const sal_uInt16 BIFF_ID_SCREENTIP          = 0x0800;
const sal_uInt16 BIFF_ID_SETUP              = 0x00A1;
const sal_uInt16 BIFF_ID_STANDARDWIDTH      = 0x0099;
const sal_uInt16 BIFF_ID_TOPMARGIN          = 0x0028;
const sal_uInt16 BIFF_ID_VCENTER            = 0x0084;
const sal_uInt16 BIFF_ID_VERPAGEBREAKS      = 0x001A;
const sal_uInt16 BIFF_ID_WSBOOL             = 0x0081;

// record constants -----------------------------------------------------------

const sal_uInt16 BIFF_COLINFO_HIDDEN        = 0x0001;
const sal_uInt16 BIFF_COLINFO_COLLAPSED     = 0x1000;

const sal_uInt16 BIFF_DEFROW_UNSYNCED       = 0x0001;
const sal_uInt16 BIFF_DEFROW_HIDDEN         = 0x0002;
const sal_uInt16 BIFF_DEFROW_SPACEABOVE     = 0x0004;
const sal_uInt16 BIFF_DEFROW_SPACEBELOW     = 0x0008;
const sal_uInt16 BIFF2_DEFROW_DEFHEIGHT     = 0x8000;
const sal_uInt16 BIFF2_DEFROW_MASK          = 0x7FFF;

const sal_uInt32 BIFF_HLINK_TARGET          = 0x00000001;   /// File name or URL.
const sal_uInt32 BIFF_HLINK_ABS             = 0x00000002;   /// Absolute path.
const sal_uInt32 BIFF_HLINK_DISPLAY         = 0x00000014;   /// Display string.
const sal_uInt32 BIFF_HLINK_LOC             = 0x00000008;   /// Target location.
const sal_uInt32 BIFF_HLINK_FRAME           = 0x00000080;   /// Target frame.
const sal_uInt32 BIFF_HLINK_UNC             = 0x00000100;   /// UNC path.

const sal_uInt16 BIFF_WSBOOL_ROWBELOW       = 0x0040;
const sal_uInt16 BIFF_WSBOOL_COLRIGHT       = 0x0080;
const sal_uInt16 BIFF_WSBOOL_FITTOPAGES     = 0x0100;

// ----------------------------------------------------------------------------

class BiffWorksheetFragment : public WorksheetHelper
{
public:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet > XSpreadsheetRef;

public:
    explicit            BiffWorksheetFragment(
                            const GlobalDataHelper& rGlobalData,
                            WorksheetType eSheetType,
                            const XSpreadsheetRef& rxSheet,
                            sal_uInt32 nSheet );

    /** Imports the entire worksheet fragment, returns true, if EOF record has been reached. */
    bool                importFragment( BiffInputStream& rStrm );

private:
    /** Imports the COLINFO record and sets column properties and formatting. */
    void                importColInfo( BiffInputStream& rStrm );
    /** Imports the BIFF2 COLUMNDEFAULT record and sets column default formatting. */
    void                importColumnDefault( BiffInputStream& rStrm );
    /** Imports the BIFF2 COLWIDTH record and sets column width. */
    void                importColWidth( BiffInputStream& rStrm );
    /** Imports the DEFCOLWIDTH record and sets default column width. */
    void                importDefColWidth( BiffInputStream& rStrm );
    /** Imports the DEFROWHEIGHT record and sets default row height and properties. */
    void                importDefRowHeight( BiffInputStream& rStrm );
    /** Imports the HLINK record and sets a cell hyperlink. */
    void                importHlink( BiffInputStream& rStrm );
    /** Imports the MEREDCELLS record and merges all cells in the document. */
    void                importMergedCells( BiffInputStream& rStrm );
    /** Imports the HORPAGEBREAKS or VERPAGEBREAKS record and inserts page breaks. */
    void                importPageBreaks( BiffInputStream& rStrm );
    /** Imports the STANDARDWIDTH record and sets standard column width. */
    void                importStandardWidth( BiffInputStream& rStrm );
    /** Imports the WSBOOL record and sets options for the current sheet. */
    void                importWsBool( BiffInputStream& rStrm );

    /** Reads and returns a string from a HLINK record. */
    ::rtl::OUString     readHlinkString( BiffInputStream& rStrm, sal_Int32 nChars, bool bUnicode );
    /** Reads and returns a string from a HLINK record. */
    ::rtl::OUString     readHlinkString( BiffInputStream& rStrm, bool bUnicode );
    /** Ignores a string in a HLINK record. */
    void                ignoreHlinkString( BiffInputStream& rStrm, sal_Int32 nChars, bool bUnicode );
    /** Ignores a string in a HLINK record. */
    void                ignoreHlinkString( BiffInputStream& rStrm, bool bUnicode );
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

