/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sheetdatacontext.hxx,v $
 *
 *  $Revision: 1.1.2.42 $
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

#ifndef OOX_XLS_SHEETDATACONTEXT_HXX
#define OOX_XLS_SHEETDATACONTEXT_HXX

#include "oox/xls/excelcontextbase.hxx"
#include "oox/xls/richstring.hxx"
#include "oox/xls/worksheethelper.hxx"

namespace com { namespace sun { namespace star {
    namespace table { class XCell; }
} } }

namespace oox {
namespace xls {

// ============================================================================

/** This class implements importing the sheetData element.

    The sheetData element contains all row settings and all cells in a single
    sheet of a spreadsheet document.
 */
class OoxSheetDataContext : public WorksheetContextBase
{
public:
    explicit            OoxSheetDataContext( const WorksheetFragmentBase& rFragment );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        onCreateContext( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    /** Imports row settings from a row element. */
    void                importRow( const ::oox::core::AttributeList& rAttribs );
    /** Imports cell settings from a c element. */
    void                importCell( const ::oox::core::AttributeList& rAttribs );
    /** Imports cell settings from an f element. */
    void                importFormula( const ::oox::core::AttributeList& rAttribs );

private:
    OoxCellData         maCurrCell;         /// Position and formatting of current imported cell.
    RichStringRef       mxInlineStr;        /// Inline rich string from 'is' element.
};

// ============================================================================

/** This class implements importing the sheetData element in external sheets.

    The sheetData element embedded in the externalBook element contains cached
    cells from externally linked sheets.
 */
class OoxExternalSheetDataContext : public WorksheetContextBase
{
public:
    explicit            OoxExternalSheetDataContext(
                            const GlobalFragmentBase& rFragment,
                            const WorksheetHelper& rSheetHelper );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    /** Imports cell settings from a c element. */
    void                importCell( const ::oox::core::AttributeList& rAttribs );

private:
    OoxCellData         maCurrCell;         /// Position and formatting of current imported cell.
};

// ============================================================================

// record identifiers ---------------------------------------------------------

const sal_uInt16 BIFF2_ID_ARRAY             = 0x0021;
const sal_uInt16 BIFF3_ID_ARRAY             = 0x0221;
const sal_uInt16 BIFF2_ID_BLANK             = 0x0001;
const sal_uInt16 BIFF3_ID_BLANK             = 0x0201;
const sal_uInt16 BIFF2_ID_BOOLERR           = 0x0005;
const sal_uInt16 BIFF3_ID_BOOLERR           = 0x0205;
const sal_uInt16 BIFF2_ID_FORMULA           = 0x0006;
const sal_uInt16 BIFF3_ID_FORMULA           = 0x0206;
const sal_uInt16 BIFF4_ID_FORMULA           = 0x0406;
const sal_uInt16 BIFF5_ID_FORMULA           = 0x0006;
const sal_uInt16 BIFF2_ID_INTEGER           = 0x0002;
const sal_uInt16 BIFF_ID_IXFE               = 0x0044;
const sal_uInt16 BIFF2_ID_LABEL             = 0x0004;
const sal_uInt16 BIFF3_ID_LABEL             = 0x0204;
const sal_uInt16 BIFF_ID_LABELSST           = 0x00FD;
const sal_uInt16 BIFF_ID_MULTBLANK          = 0x00BE;
const sal_uInt16 BIFF_ID_MULTRK             = 0x00BD;
const sal_uInt16 BIFF2_ID_NUMBER            = 0x0003;
const sal_uInt16 BIFF3_ID_NUMBER            = 0x0203;
const sal_uInt16 BIFF_ID_RK                 = 0x027E;
const sal_uInt16 BIFF2_ID_ROW               = 0x0008;
const sal_uInt16 BIFF3_ID_ROW               = 0x0208;
const sal_uInt16 BIFF_ID_RSTRING            = 0x00D6;
const sal_uInt16 BIFF_ID_SHRFMLA            = 0x04BC;
const sal_uInt16 BIFF2_ID_STRING            = 0x0007;
const sal_uInt16 BIFF3_ID_STRING            = 0x0207;

// record constants -----------------------------------------------------------

const sal_Int32 BIFF_IXFE_USE_CACHED        = 63;

const sal_uInt8 BIFF2_XF_MASK               = 0x3F;

const sal_uInt8 BIFF_BOOLERR_BOOL           = 0;
const sal_uInt8 BIFF_BOOLERR_ERROR          = 1;

const sal_uInt8 BIFF_FORMULA_RES_STRING     = 0x00;    /// Result is a string.
const sal_uInt8 BIFF_FORMULA_RES_BOOL       = 0x01;    /// Result is Boolean value.
const sal_uInt8 BIFF_FORMULA_RES_ERROR      = 0x02;    /// Result is error code.
const sal_uInt8 BIFF_FORMULA_RES_EMPTY      = 0x03;    /// Result is empty cell (BIFF8 only).

const sal_uInt16 BIFF_FORMULA_SHARED        = 0x0008;  /// Shared formula cell.

const sal_uInt8 BIFF2_ROW_FORMATTED         = 0x01;
const sal_uInt16 BIFF_ROW_COLLAPSED         = 0x0010;
const sal_uInt16 BIFF_ROW_HIDDEN            = 0x0020;
const sal_uInt16 BIFF_ROW_UNSYNCED          = 0x0040;
const sal_uInt16 BIFF_ROW_FORMATTED         = 0x0080;
const sal_uInt16 BIFF_ROW_XFMASK            = 0x0FFF;
const sal_uInt16 BIFF_ROW_FLAGDEFHEIGHT     = 0x8000;
const sal_uInt16 BIFF_ROW_HEIGHTMASK        = 0x7FFF;

// ----------------------------------------------------------------------------

/** This class implements importing row settings and all cells of a sheet.
 */
class BiffSheetDataContext : public WorksheetHelper
{
public:
    explicit            BiffSheetDataContext( const WorksheetHelper& rSheetHelper );

    /** Tries to import a sheet data record. */
    void                importRecord( BiffInputStream& rStrm );

private:
    /** Sets current cell according to the passed address. */
    void                setCurrCell( const BiffAddress& rBiffAddr );

    /** Imports an XF identifier and sets the mnXfId member. */
    void                importXfId( BiffInputStream& rStrm, bool bBiff2 );
    /** Imports a BIFF cell address and the following XF identifier. */
    void                importCellHeader( BiffInputStream& rStrm, bool bBiff2 );

    /** Imports a BLANK record describing a blank but formatted cell. */
    void                importBlank( BiffInputStream& rStrm );
    /** Imports a BOOLERR record describing a boolean or error code cell. */
    void                importBoolErr( BiffInputStream& rStrm );
    /** Imports a FORMULA record describing a formula cell. */
    void                importFormula( BiffInputStream& rStrm );
    /** Imports an INTEGER record describing a BIFF2 integer cell. */
    void                importInteger( BiffInputStream& rStrm );
    /** Imports a LABEL record describing an unformatted string cell. */
    void                importLabel( BiffInputStream& rStrm );
    /** Imports a LABELSST record describing a string cell using the shared string list. */
    void                importLabelSst( BiffInputStream& rStrm );
    /** Imports a MULTBLANK record describing a range of blank but formatted cells. */
    void                importMultBlank( BiffInputStream& rStrm );
    /** Imports a MULTRK record describing a range of numeric cells. */
    void                importMultRk( BiffInputStream& rStrm );
    /** Imports a NUMBER record describing a floating-point cell. */
    void                importNumber( BiffInputStream& rStrm );
    /** Imports an RK record describing a numeric cell. */
    void                importRk( BiffInputStream& rStrm );

    /** Imports row settings from a ROW record. */
    void                importRow( BiffInputStream& rStrm );
    /** Imports an ARRAY record describing an array formula of a cell range. */
    void                importArray( BiffInputStream& rStrm );

private:
    OoxCellData         maCurrCell;             /// Position and formatting of current imported cell.
    sal_uInt32          mnFormulaIgnoreSize;    /// Number of bytes to be ignored in FORMULA record.
    sal_uInt32          mnArrayIgnoreSize;      /// Number of bytes to be ignored in ARRAY record.
    sal_uInt16          mnBiff2XfId;            /// Current XF identifier from IXFE record.
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

