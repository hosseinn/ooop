/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: worksheethelper.cxx,v $
 *
 *  $Revision: 1.1.2.44 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:49 $
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

#include "oox/xls/worksheethelper.hxx"
#include <utility>
#include <list>
#include <rtl/ustrbuf.hxx>
#include <com/sun/star/container/XNameContainer.hpp>
#include <com/sun/star/util/XMergeable.hpp>
#include <com/sun/star/table/XColumnRowRange.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/sheet/XCellAddressable.hpp>
#include <com/sun/star/sheet/XCellRangeAddressable.hpp>
#include <com/sun/star/sheet/XFormulaTokens.hpp>
#include <com/sun/star/sheet/XSheetOutline.hpp>
#include <com/sun/star/sheet/XNamedRange.hpp>
#include <com/sun/star/style/XStyleFamiliesSupplier.hpp>
#include <com/sun/star/style/XStyle.hpp>
#include <com/sun/star/text/XText.hpp>
#include "tokens.hxx"
#include "oox/core/containerhelper.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/defnamesbuffer.hxx"
#include "oox/xls/formulaparser.hxx"
#include "oox/xls/pagestyle.hxx"
#include "oox/xls/richstring.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/stylesbuffer.hxx"
#include "oox/xls/unitconverter.hxx"
#include "oox/xls/viewsettings.hxx"
#include "oox/xls/worksheetbuffer.hxx"

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::container::XIndexAccess;
using ::com::sun::star::container::XNamed;
using ::com::sun::star::container::XNameAccess;
using ::com::sun::star::container::XNameContainer;
using ::com::sun::star::util::XMergeable;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::table::BorderLine;
using ::com::sun::star::table::XColumnRowRange;
using ::com::sun::star::table::XTableColumns;
using ::com::sun::star::table::XTableRows;
using ::com::sun::star::table::XCell;
using ::com::sun::star::table::XCellRange;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::sheet::XCellAddressable;
using ::com::sun::star::sheet::XCellRangeAddressable;
using ::com::sun::star::sheet::XFormulaTokens;
using ::com::sun::star::sheet::XNamedRange;
using ::com::sun::star::sheet::XSheetOutline;
using ::com::sun::star::style::XStyleFamiliesSupplier;
using ::com::sun::star::style::XStyle;
using ::com::sun::star::text::XText;
using ::com::sun::star::text::XTextContent;
using ::com::sun::star::text::XTextRange;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

void OoxCellData::reset()
{
    mxCell.clear();
    maValueStr = maFormulaStr = maFormulaRef = OUString();
    mnCellType = mnFormulaType = XML_TOKEN_INVALID;
    mnSharedId = mnXfId = mnNumFmtId = -1;
    mbHasValueStr = false;
}

// ----------------------------------------------------------------------------

OoxColumnData::OoxColumnData() :
    mnFirstCol( -1 ),
    mnLastCol( -1 ),
    mfWidth( 0.0 ),
    mnXfId( -1 ),
    mnLevel( 0 ),
    mbHidden( false ),
    mbCollapsed( false )
{
}

bool OoxColumnData::tryExpand( const OoxColumnData& rNewData )
{
    bool bExpandable =
        (mnFirstCol <= rNewData.mnFirstCol) &&
        (rNewData.mnFirstCol <= mnLastCol + 1) &&
        (mfWidth == rNewData.mfWidth) &&
        // ignore mnXfId, cell formatting is always set directly
        (mnLevel == rNewData.mnLevel) &&
        (mbHidden == rNewData.mbHidden) &&
        (mbCollapsed == rNewData.mbCollapsed);
    if( bExpandable )
        mnLastCol = rNewData.mnLastCol;
    return bExpandable;
}

// ----------------------------------------------------------------------------

OoxRowData::OoxRowData() :
    mnFirstRow( -1 ),
    mnLastRow( -1 ),
    mfHeight( 0.0 ),
    mnXfId( -1 ),
    mnLevel( 0 ),
    mbFormatted( false ),
    mbHidden( false ),
    mbCollapsed( false )
{
}

bool OoxRowData::tryExpand( const OoxRowData& rNewData )
{
    bool bExpandable =
        (mnFirstRow <= rNewData.mnFirstRow) &&
        (rNewData.mnFirstRow <= mnLastRow + 1) &&
        (mfHeight == rNewData.mfHeight) &&
        // ignore mnXfId, cell formatting is always set directly
        (mnLevel == rNewData.mnLevel) &&
        (mbHidden == rNewData.mbHidden) &&
        (mbCollapsed == rNewData.mbCollapsed);
    if( bExpandable )
        mnLastRow = rNewData.mnLastRow;
    return bExpandable;
}

// ----------------------------------------------------------------------------

OoxPageBreakData::OoxPageBreakData() :
    mnColRow( 0 ),
    mbManual( false )
{
}

// ----------------------------------------------------------------------------

OoxHyperlinkData::OoxHyperlinkData()
{
}

// ============================================================================
// ============================================================================

SharedFormulaBuffer::SharedFormulaBuffer( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void SharedFormulaBuffer::importSharedFmla( const OUString& rFormula, sal_Int32 nSharedId, const CellAddress& rBaseAddr )
{
    // unique key for the shared formula from base address
    OSL_ENSURE( maIndexMap.count( nSharedId ) == 0, "SharedFormulaBuffer::importSharedFmla - shared formula exists already" );

    // create the defined name representing the shared formula
    Reference< XNamedRange > xNamedRange = createDefinedName( rBaseAddr.Sheet, nSharedId );

    // convert the formula definition
    Reference< XFormulaTokens > xTokens( xNamedRange, UNO_QUERY );
    if( xTokens.is() )
    {
        SimpleFormulaContext aContext( xTokens );
        aContext.setBaseAddress( rBaseAddr, true );
        getFormulaParser().importFormula( aContext, rFormula );
    }
}

void SharedFormulaBuffer::importShrFmla( BiffInputStream& rStrm, const CellAddress& rBaseAddr )
{
    BiffRange aBiffRange;
    aBiffRange.read( rStrm, false );        // always 8bit column indexes
    CellRangeAddress aRange;
    if( getAddressConverter().convertToCellRange( aRange, aBiffRange, rBaseAddr.Sheet, true ) )
    {
        OSL_ENSURE( (aRange.StartColumn <= rBaseAddr.Column) && (rBaseAddr.Column <= aRange.EndColumn) &&
            (aRange.StartRow <= rBaseAddr.Row) && (rBaseAddr.Row <= aRange.EndRow),
            "SharedFormulaBuffer::importShrFmla - invalid range for shared formula" );

        // unique key for the shared formula from base address
        sal_Int32 nSharedId = AddressConverter::convertToId( rBaseAddr );
        OSL_ENSURE( maIndexMap.count( nSharedId ) == 0, "SharedFormulaBuffer::importShrFmla - shared formula exists already" );

        // create the defined name representing the shared formula
        Reference< XNamedRange > xNamedRange = createDefinedName( rBaseAddr.Sheet, nSharedId );

        // load the formula definition
        Reference< XFormulaTokens > xTokens( xNamedRange, UNO_QUERY );
        if( xTokens.is() )
        {
            rStrm.ignore( 2 );
            SimpleFormulaContext aContext( xTokens );
            aContext.setBaseAddress( rBaseAddr, true );
            getFormulaParser().importFormula( aContext, rStrm );
        }
    }
}

sal_Int32 SharedFormulaBuffer::getTokenIndexFromId( sal_Int32 nSharedId ) const
{
    TokenIndexMap::const_iterator aIt = maIndexMap.find( nSharedId );
    return (aIt == maIndexMap.end()) ? -1 : aIt->second;
}

Reference< XNamedRange > SharedFormulaBuffer::createDefinedName( sal_Int32 nSheet, sal_Int32 nSharedId )
{
    // create the defined name representing the shared formula
    OUString aName = OUStringBuffer().appendAscii( "__shared_" ).
        append( nSheet + 1 ).append( sal_Unicode( '_' ) ).append( nSharedId ).makeStringAndClear();
    Reference< XNamedRange > xNamedRange = getDefinedNames().createDefinedName( aName );
    sal_Int32 nTokenIndex = getDefinedNames().getTokenIndex( xNamedRange );
    if( nTokenIndex >= 0 )
        maIndexMap[ nSharedId ] = nTokenIndex;
    return xNamedRange;
}

// ============================================================================

class WorksheetHelperImpl : public GlobalDataHelper
{
public:
    explicit            WorksheetHelperImpl(
                            const GlobalDataHelper& rGlobalData,
                            WorksheetType eSheetType,
                            const Reference< XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );

    /** Returns a cell formula simulating the passed boolean value. */
    const OUString&     getBooleanFormula( bool bValue ) const;
    /** Returns a BIFF error code from the passed error string. */
    sal_uInt8           getBiffErrorCode( const OUString& rErrorCode ) const;

    /** Returns the type of this sheet. */
    inline WorksheetType getSheetType() const { return meSheetType; }
    /** Returns the XSpreadsheet interface of the processed sheet. */
    inline Reference< XSpreadsheet > getXSpreadsheet() const { return mxSheet; }
    /** Returns the index of the current sheet. */
    inline sal_Int16    getSheetIndex() const { return mnSheet; }

    /** Increment the conditional formatting index by one, and returns the old
        index. */
    inline sal_Int32    incCondFormatIndex() { return mnCondFormatId++; }
    /** Returns the index of the current conditional formatting item. */
    inline sal_Int32    getCondFormatIndex() const { return mnCondFormatId; }

    /** Returns the XCell interface for the passed cell address. */
    Reference< XCell >  getCell( const CellAddress& rAddress ) const;
    /** Returns the XCellRange interface for the passed cell range address. */
    Reference< XCellRange > getCellRange( const CellRangeAddress& rRange ) const;

    /** Returns the XCellRange interface for a column. */
    Reference< XCellRange > getColumn( sal_Int32 nCol ) const;
    /** Returns the XCellRange interface for a row. */
    Reference< XCellRange > getRow( sal_Int32 nRow ) const;

    /** Returns the XTableColumns interface for a range of columns. */
    Reference< XTableColumns > getColumns( sal_Int32 nFirstCol, sal_Int32 nLastCol ) const;
    /** Returns the XTableRows interface for a range of rows. */
    Reference< XTableRows > getRows( sal_Int32 nFirstRow, sal_Int32 nLastRow ) const;

    /** Returns the buffer containing all shared formulas in this sheet. */
    inline SharedFormulaBuffer& getSharedFormulas() { return maSharedFmlas; }
    /** Returns the page style object that implements import/export of page/print settings. */
    inline PageStyle&   getPageStyle() { return maPageStyle; }
    /** Returns the global phonetic settings for this sheet. */
    inline PhoneticSettings& getPhoneticSettings() { return maPhoneticSett; }

    /** Stores the cell format at the passed address. */
    void                setCellFormat( const OoxCellData& rCellData );
    /** Inserts the hyperlink URL into the passed cell range. */
    void                setHyperlink( const OoxHyperlinkData& rHyperlink );
    /** Merges the cells in the passed cell range. */
    void                setMergedRange( const CellRangeAddress& rRange );

    /** Sets base width for all columns (without padding pixels). This value
        is only used, if base width has not been set with setDefaultColumnWidth(). */
    void                setBaseColumnWidth( sal_Int32 nWidth );
    /** Sets default width for all columns. This function overrides the base
        width set with the setBaseColumnWidth() function. */
    void                setDefaultColumnWidth( double fWidth );
    /** Sets column settings for a specific column range.
        @descr  Column default formatting is converted directly, other settings
        are cached and converted in the finalizeImport() call. */
    void                setColumnData( const OoxColumnData& rData );

    /** Sets default height and hidden state for all unused rows in the sheet. */
    void                setDefaultRowSettings( double fHeight, bool bHidden );
    /** Sets row settings for a specific row.
        @descr  Row default formatting is converted directly, other settings
        are cached and converted in the finalizeImport() call. */
    void                setRowData( const OoxRowData& rData );

    /** Sets the position of outline summary symbols for this sheet. */
    void                setOutlineSummarySymbols( bool bSummaryRight, bool bSummaryBelow );

    /** Converts column default cell formatting. */
    void                convertColumnFormat( sal_Int32 nFirstCol, sal_Int32 nLastCol, sal_Int32 nXfId );
    /** Converts row default cell formatting. */
    void                convertRowFormat( sal_Int32 nFirstRow, sal_Int32 nLastRow, sal_Int32 nXfId );
    /** Converts a column or row page break described in the passed struct. */
    void                convertPageBreak( const OoxPageBreakData& rData, bool bRowBreak );

    /** Final conversion after importing the worksheet. */
    void                finalizeWorksheetImport();

private:
    typedef ::std::vector< sal_Int32 >              OutlineLevelVec;
    typedef ::std::map< OUString, sal_uInt8 >       ErrorCodeMap;
    typedef ::std::map< sal_Int32, OoxColumnData >  OoxColumnDataMap;
    typedef ::std::map< sal_Int32, OoxRowData >     OoxRowDataMap;
    typedef ::std::list< OoxHyperlinkData >         OoxHyperlinkList;

    struct XfIdRange
    {
        CellRangeAddress    maRange;            /// The formatted cell range.
        sal_Int32           mnXfId;             /// XF identifier for the range.
        sal_Int32           mnNumFmtId;         /// Number format id overriding the XF.

        void                set( const OoxCellData& rCellData );
        bool                tryExpand( const OoxCellData& rCellData );
        bool                tryMerge( const XfIdRange& rXfIdRange );
    };

    struct MergedRange
    {
        CellRangeAddress    maRange;            /// The formatted cell range.
        sal_Int32           mnHorAlign;         /// Horizontal alignment in the range.

        explicit            MergedRange( const CellRangeAddress& rAddress );
        explicit            MergedRange( const CellAddress& rAddress, sal_Int32 nHorAlign );
        bool                tryExpand( const CellAddress& rAddress, sal_Int32 nHorAlign );
    };

    typedef ::std::pair< sal_Int32, sal_Int32 > RowColKey;
    typedef ::std::map< RowColKey, XfIdRange >  XfIdRangeMap;
    typedef ::std::list< MergedRange >          MergedRangeList;

    /** Writes all cell formatting attributes to the passed cell range. */
    void                writeXfIdRangeProperties( const XfIdRange& rXfIdRange ) const;
    /** Tries to merge the ranges last inserted in maXfIdRanges with existing ranges. */
    void                mergeXfIdRanges();
    /** Finalizes the remaining ranges in maXfIdRanges. */
    void                finalizeXfIdRanges();

    /** Inserts all imported hyperlinks into their cell ranges. */
    void                finalizeHyperlinkRanges() const;
    /** Inserts a hyperlinks into the specified cell. */
    void                finalizeHyperlink( const CellAddress& rAddress, const OUString& rUrl ) const;

    /** Merges all cached merged ranges and updates right/bottom cell borders. */
    void                finalizeMergedRanges() const;
    /** Merges the passed merged range and updates right/bottom cell borders. */
    void                finalizeMergedRange( const CellRangeAddress& rRange ) const;

    /** Converts imported page style properties and inserts a page style into current sheet. */
    void                convertPageStyle();

    /** Converts column properties for all columns in the sheet. */
    void                convertColumns();
    /** Converts column properties. */
    void                convertColumns( OutlineLevelVec& orColLevels, sal_Int32 nFirstCol, sal_Int32 nLastCol, const OoxColumnData& rData );

    /** Converts row properties for all rows in the sheet. */
    void                convertRows();
    /** Converts row properties. */
    void                convertRows( OutlineLevelVec& orRowLevels, sal_Int32 nFirstRow, sal_Int32 nLastRow, const OoxRowData& rData );

    /** Converts outline grouping for the passed column or row. */
    void                convertOutlines( OutlineLevelVec& orLevels, sal_Int32 nColRow, sal_Int32 nLevel, bool bCollapsed, bool bRows );
    /** Groups columns or rows for the given range. */
    void                groupColumnsOrRows( sal_Int32 nFirstColRow, sal_Int32 nLastColRow, bool bCollapsed, bool bRows );

private:
    const OUString      maTrueFormula;      /// Replacement formula for TRUE boolean cells.
    const OUString      maFalseFormula;     /// Replacement formula for FALSE boolean cells.
    const OUString      maRightBorderProp;  /// Property name of the right border of a cell.
    const OUString      maBottomBorderProp; /// Property name of the bottom border of a cell.
    const OUString      maWidthProp;        /// Property name for column width.
    const OUString      maHeightProp;       /// Property name for row height.
    const OUString      maVisibleProp;      /// Property name for column/row visibility.
    const OUString      maPageBreakProp;    /// Property name of a page break.
    const OUString      maUrlTextField;     /// Service name for a URL text field.
    const OUString      maUrlProp;          /// Property name for the URL string in a URL text field.
    const OUString      maReprProp;         /// Property name for the URL representation in a URL text field.
    const CellAddress&  mrMaxApiPos;        /// Reference to maximum Calc cell address from address converter.
    ErrorCodeMap        maErrorCodes;       /// Maps error code strings to BIFF error constants.
    OoxColumnData       maDefColData;       /// Default column formatting.
    OoxColumnDataMap    maColDatas;         /// Column data sorted by first column index.
    OoxRowData          maDefRowData;       /// Default row formatting.
    OoxRowDataMap       maRowDatas;         /// Row data sorted by row index.
    OoxHyperlinkList    maHyperlinks;       /// Cell ranges containing hyperlinks.
    XfIdRangeMap        maXfIdRanges;       /// Collected XF identifiers for cell ranges.
    MergedRangeList     maMergedRanges;     /// Merged cell ranges.
    MergedRangeList     maCenterFillRanges; /// Merged cell ranges from 'center across' or 'fill' alignment.
    SharedFormulaBuffer maSharedFmlas;      /// Buffer for shared formulas in this sheet.
    PageStyle           maPageStyle;        /// Page style object for this sheet.
    PhoneticSettings    maPhoneticSett;     /// Global phonetic settings for this sheet.
    const WorksheetType meSheetType;        /// Type of thes sheet.
    Reference< XSpreadsheet > mxSheet;      /// Reference to the current sheet.
    const sal_Int16     mnSheet;            /// Index of the current sheet.
    sal_Int32           mnCondFormatId;     /// Index of the conditional formatting item per sheet.
    bool                mbHasDefWidth;      /// True = default column width is set from defaultColWidth attribute.
    bool                mbSummaryRight;     /// True = column outline summary symbol right of group.
    bool                mbSummaryBelow;     /// True = row outline summary symbol below group.
};

// ----------------------------------------------------------------------------

WorksheetHelperImpl::WorksheetHelperImpl( const GlobalDataHelper& rGlobalData,
        WorksheetType eSheetType, const Reference< XSpreadsheet >& rxSheet, sal_Int32 nSheet ) :
    GlobalDataHelper( rGlobalData ),
    maTrueFormula( CREATE_OUSTRING( "=TRUE()" ) ),
    maFalseFormula( CREATE_OUSTRING( "=FALSE()" ) ),
    maRightBorderProp( CREATE_OUSTRING( "RightBorder" ) ),
    maBottomBorderProp( CREATE_OUSTRING( "BottomBorder" ) ),
    maWidthProp( CREATE_OUSTRING( "Width" ) ),
    maHeightProp( CREATE_OUSTRING( "Height" ) ),
    maVisibleProp( CREATE_OUSTRING( "IsVisible" ) ),
    maPageBreakProp( CREATE_OUSTRING( "IsStartOfNewPage" ) ),
    maUrlTextField( CREATE_OUSTRING( "com.sun.star.text.TextField.URL" ) ),
    maUrlProp( CREATE_OUSTRING( "URL" ) ),
    maReprProp( CREATE_OUSTRING( "Representation" ) ),
    mrMaxApiPos( rGlobalData.getAddressConverter().getMaxApiAddress() ),
    maSharedFmlas( rGlobalData ),
    maPageStyle( rGlobalData ),
    maPhoneticSett( rGlobalData ),
    meSheetType( eSheetType ),
    mxSheet( rxSheet ),
    mnSheet( static_cast< sal_Int16 >( nSheet ) ),
    mnCondFormatId( 0 ),
    mbHasDefWidth( false ),
    mbSummaryRight( true ),
    mbSummaryBelow( true )
{
    OSL_ENSURE( rxSheet.is(), "WorksheetHelperImpl::WorksheetHelperImpl - missing XSpreadsheet interface" );
    OSL_ENSURE( nSheet <= SAL_MAX_INT16, "WorksheetHelperImpl::WorksheetHelperImpl - invalid sheet index" );

    // map error code names to BIFF error codes
    maErrorCodes[ CREATE_OUSTRING( "#NULL!" ) ]  = BIFF_ERR_NULL;
    maErrorCodes[ CREATE_OUSTRING( "#DIV/0!" ) ] = BIFF_ERR_DIV0;
    maErrorCodes[ CREATE_OUSTRING( "#VALUE!" ) ] = BIFF_ERR_VALUE;
    maErrorCodes[ CREATE_OUSTRING( "#REF!" ) ]   = BIFF_ERR_REF;
    maErrorCodes[ CREATE_OUSTRING( "#NAME?" ) ]  = BIFF_ERR_NAME;
    maErrorCodes[ CREATE_OUSTRING( "#NUM!" ) ]   = BIFF_ERR_NUM;
    maErrorCodes[ CREATE_OUSTRING( "#NA" ) ]     = BIFF_ERR_NA;

    // default column settings (width and hidden state may be updated later)
    maDefColData.mfWidth = 8.5;
    maDefColData.mnXfId = -1;
    maDefColData.mnLevel = 0;
    maDefColData.mbHidden = false;
    maDefColData.mbCollapsed = false;

    // default row settings (height and hidden state may be updated later)
    maDefRowData.mfHeight = 0.0;
    maDefRowData.mnXfId = -1;
    maDefRowData.mnLevel = 0;
    maDefRowData.mbFormatted = false;
    maDefRowData.mbHidden = false;
    maDefRowData.mbCollapsed = false;
}

const OUString& WorksheetHelperImpl::getBooleanFormula( bool bValue ) const
{
    return bValue ? maTrueFormula : maFalseFormula;
}

sal_uInt8 WorksheetHelperImpl::getBiffErrorCode( const OUString& rErrorCode ) const
{
    ErrorCodeMap::const_iterator aIt = maErrorCodes.find( rErrorCode );
    return (aIt == maErrorCodes.end()) ? BIFF_ERR_NA : aIt->second;
}

Reference< XCell > WorksheetHelperImpl::getCell( const CellAddress& rAddress ) const
{
    Reference< XCell > xCell;
    try
    {
        xCell = mxSheet->getCellByPosition( rAddress.Column, rAddress.Row );
    }
    catch( Exception& )
    {
    }
    return xCell;
}

Reference< XCellRange > WorksheetHelperImpl::getCellRange( const CellRangeAddress& rRange ) const
{
    Reference< XCellRange > xRange;
    try
    {
        xRange = mxSheet->getCellRangeByPosition( rRange.StartColumn, rRange.StartRow, rRange.EndColumn, rRange.EndRow );
    }
    catch( Exception& )
    {
    }
    return xRange;
}

Reference< XCellRange > WorksheetHelperImpl::getColumn( sal_Int32 nCol ) const
{
    Reference< XCellRange > xColumn;
    try
    {
        Reference< XColumnRowRange > xColRowRange( mxSheet, UNO_QUERY_THROW );
        Reference< XTableColumns > xColumns = xColRowRange->getColumns();
        if( xColumns.is() )
            xColumn.set( xColumns->getByIndex( nCol ), UNO_QUERY );
    }
    catch( Exception& )
    {
    }
    return xColumn;
}

Reference< XCellRange > WorksheetHelperImpl::getRow( sal_Int32 nRow ) const
{
    Reference< XCellRange > xRow;
    try
    {
        Reference< XColumnRowRange > xColRowRange( mxSheet, UNO_QUERY_THROW );
        Reference< XTableRows > xRows = xColRowRange->getRows();
            xRow.set( xRows->getByIndex( nRow ), UNO_QUERY );
    }
    catch( Exception& )
    {
    }
    return xRow;
}

Reference< XTableColumns > WorksheetHelperImpl::getColumns( sal_Int32 nFirstCol, sal_Int32 nLastCol ) const
{
    Reference< XTableColumns > xColumns;
    nLastCol = ::std::min( nLastCol, mrMaxApiPos.Column );
    if( (0 <= nFirstCol) && (nFirstCol <= nLastCol) )
    {
        Reference< XColumnRowRange > xRange( getCellRange( CellRangeAddress( mnSheet, nFirstCol, 0, nLastCol, 0 ) ), UNO_QUERY );
        if( xRange.is() )
            xColumns = xRange->getColumns();
    }
    return xColumns;
}

Reference< XTableRows > WorksheetHelperImpl::getRows( sal_Int32 nFirstRow, sal_Int32 nLastRow ) const
{
    Reference< XTableRows > xRows;
    nLastRow = ::std::min( nLastRow, mrMaxApiPos.Row );
    if( (0 <= nFirstRow) && (nFirstRow <= nLastRow) )
    {
        Reference< XColumnRowRange > xRange( getCellRange( CellRangeAddress( mnSheet, 0, nFirstRow, 0, nLastRow ) ), UNO_QUERY );
        if( xRange.is() )
            xRows = xRange->getRows();
    }
    return xRows;
}

void WorksheetHelperImpl::setCellFormat( const OoxCellData& rCellData )
{
    if( rCellData.mxCell.is() && (rCellData.mnXfId >= 0) || (rCellData.mnNumFmtId >= 0) )
    {
        // try to merge existing ranges and to write some formatting properties
        if( !maXfIdRanges.empty() )
        {
            // get row index of last inserted cell
            sal_Int32 nLastRow = maXfIdRanges.rbegin()->second.maRange.StartRow;
            // row changed - try to merge ranges of last row with existing ranges
            if( rCellData.maAddress.Row != nLastRow )
            {
                mergeXfIdRanges();
                // write format properties of all ranges above last row and remove them
                XfIdRangeMap::iterator aIt = maXfIdRanges.begin(), aEnd = maXfIdRanges.end();
                while( aIt != aEnd )
                {
                    if( aIt->second.maRange.EndRow < nLastRow )
                    {
                        writeXfIdRangeProperties( aIt->second );
                        maXfIdRanges.erase( aIt++ );
                    }
                    else
                        ++aIt;
                }
            }
        }

        // try to expand last existing range, or create new range entry
        if( maXfIdRanges.empty() || !maXfIdRanges.rbegin()->second.tryExpand( rCellData ) )
            maXfIdRanges[ RowColKey( rCellData.maAddress.Row, rCellData.maAddress.Column ) ].set( rCellData );

        // update merged ranges for 'center across selection' and 'fill'
        if( const Xf* pXf = getStyles().getCellXf( rCellData.mnXfId ).get() )
        {
            sal_Int32 nHorAlign = pXf->getAlignment().getOoxData().mnHorAlign;
            if( (nHorAlign == XML_centerContinuous) || (nHorAlign == XML_fill) )
            {
                /*  start new merged range, if cell is not empty (#108781#),
                    or try to expand last range with empty cell */
                if( rCellData.mnCellType != XML_TOKEN_INVALID )
                    maCenterFillRanges.push_back( MergedRange( rCellData.maAddress, nHorAlign ) );
                else if( !maCenterFillRanges.empty() )
                    maCenterFillRanges.rbegin()->tryExpand( rCellData.maAddress, nHorAlign );
            }
        }
    }
}

void WorksheetHelperImpl::setHyperlink( const OoxHyperlinkData& rHyperlink )
{
    maHyperlinks.push_back( rHyperlink );
}

void WorksheetHelperImpl::setMergedRange( const CellRangeAddress& rRange )
{
    maMergedRanges.push_back( MergedRange( rRange ) );
}

void WorksheetHelperImpl::setBaseColumnWidth( sal_Int32 nWidth )
{
    // do not modify width, if setDefaultColumnWidth() has been used
    if( !mbHasDefWidth && (nWidth > 0) )
    {
        /*  #i3006# add 5 pixels padding to the width, assuming 1 pixel =
            1/96 inch. => 5/96 inch == 1.32 mm. */
        const UnitConverter& rUnitConv = getUnitConverter();
        maDefColData.mfWidth = rUnitConv.calcDigitsFromMm100( rUnitConv.calcMm100FromDigits( nWidth ) + 132 );
    }
}

void WorksheetHelperImpl::setDefaultColumnWidth( double fWidth )
{
    // overrides a width set with setBaseColumnWidth()
    if( fWidth > 0.0 )
    {
        maDefColData.mfWidth = fWidth;
        mbHasDefWidth = true;
    }
}

void WorksheetHelperImpl::setColumnData( const OoxColumnData& rData )
{
    // convert 1-based OOX column indexes to 0-based API column indexes
    sal_Int32 nFirstCol = rData.mnFirstCol - 1;
    sal_Int32 nLastCol = rData.mnLastCol - 1;
    if( nFirstCol <= mrMaxApiPos.Column )
    {
        // set column formatting directly, nLastCol is checked inside the function
        convertColumnFormat( nFirstCol, nLastCol, rData.mnXfId );
        // expand last entry or add new entry
        if( maColDatas.empty() || !maColDatas.rbegin()->second.tryExpand( rData ) )
            maColDatas[ nFirstCol ] = rData;
    }
}

void WorksheetHelperImpl::setDefaultRowSettings( double fHeight, bool bHidden )
{
    maDefRowData.mfHeight = fHeight;
    maDefRowData.mbHidden = bHidden;
}

void WorksheetHelperImpl::setRowData( const OoxRowData& rData )
{
    // convert 1-based OOX row indexes to 0-based API row indexes
    sal_Int32 nFirstRow = rData.mnFirstRow - 1;
    sal_Int32 nLastRow = rData.mnLastRow - 1;
    if( nFirstRow <= mrMaxApiPos.Row )
    {
        // set row formatting directly
        if( rData.mbFormatted )
            convertRowFormat( nFirstRow, nLastRow, rData.mnXfId );
        // expand last entry or add new entry
        if( maRowDatas.empty() || !maRowDatas.rbegin()->second.tryExpand( rData ) )
            maRowDatas[ nFirstRow ] = rData;
    }
}

void WorksheetHelperImpl::setOutlineSummarySymbols( bool bSummaryRight, bool bSummaryBelow )
{
    mbSummaryRight = bSummaryRight;
    mbSummaryBelow = bSummaryBelow;
}

void WorksheetHelperImpl::convertColumnFormat( sal_Int32 nFirstCol, sal_Int32 nLastCol, sal_Int32 nXfId )
{
    CellRangeAddress aRange( mnSheet, nFirstCol, 0, nLastCol, mrMaxApiPos.Row );
    if( getAddressConverter().validateCellRange( aRange, false ) )
    {
        PropertySet aPropSet( getCellRange( aRange ) );
        getStyles().writeCellXfToPropertySet( aPropSet, nXfId );
    }
}

void WorksheetHelperImpl::convertRowFormat( sal_Int32 nFirstRow, sal_Int32 nLastRow, sal_Int32 nXfId )
{
    CellRangeAddress aRange( mnSheet, 0, nFirstRow, mrMaxApiPos.Column, nLastRow );
    if( getAddressConverter().validateCellRange( aRange, false ) )
    {
        PropertySet aPropSet( getCellRange( aRange ) );
        getStyles().writeCellXfToPropertySet( aPropSet, nXfId );
    }
}

void WorksheetHelperImpl::convertPageBreak( const OoxPageBreakData& rData, bool bRowBreak )
{
    if( rData.mbManual && (rData.mnColRow > 0) )
    {
        PropertySet aPropSet( bRowBreak ? getRow( rData.mnColRow ) : getColumn( rData.mnColRow ) );
        aPropSet.setProperty( maPageBreakProp, true );
    }
}

void WorksheetHelperImpl::finalizeWorksheetImport()
{
    finalizeXfIdRanges();
    finalizeHyperlinkRanges();
    finalizeMergedRanges();
    convertPageStyle();
    convertColumns();
    convertRows();
}

// private --------------------------------------------------------------------

void WorksheetHelperImpl::XfIdRange::set( const OoxCellData& rCellData )
{
    maRange.Sheet = rCellData.maAddress.Sheet;
    maRange.StartColumn = maRange.EndColumn = rCellData.maAddress.Column;
    maRange.StartRow = maRange.EndRow = rCellData.maAddress.Row;
    mnXfId = rCellData.mnXfId;
    mnNumFmtId = rCellData.mnNumFmtId;
}

bool WorksheetHelperImpl::XfIdRange::tryExpand( const OoxCellData& rCellData )
{
    if( (mnXfId == rCellData.mnXfId) && (mnNumFmtId == rCellData.mnNumFmtId) &&
        (maRange.StartRow == rCellData.maAddress.Row) &&
        (maRange.EndRow == rCellData.maAddress.Row) &&
        (maRange.EndColumn + 1 == rCellData.maAddress.Column) )
    {
        ++maRange.EndColumn;
        return true;
    }
    return false;
}

bool WorksheetHelperImpl::XfIdRange::tryMerge( const XfIdRange& rXfIdRange )
{
    if( (mnXfId == rXfIdRange.mnXfId) &&
        (mnNumFmtId == rXfIdRange.mnNumFmtId) &&
        (maRange.EndRow + 1 == rXfIdRange.maRange.StartRow) &&
        (maRange.StartColumn == rXfIdRange.maRange.StartColumn) &&
        (maRange.EndColumn == rXfIdRange.maRange.EndColumn) )
    {
        maRange.EndRow = rXfIdRange.maRange.EndRow;
        return true;
    }
    return false;
}


WorksheetHelperImpl::MergedRange::MergedRange( const CellRangeAddress& rRange ) :
    maRange( rRange ),
    mnHorAlign( XML_TOKEN_INVALID )
{
}

WorksheetHelperImpl::MergedRange::MergedRange( const CellAddress& rAddress, sal_Int32 nHorAlign ) :
    maRange( rAddress.Sheet, rAddress.Column, rAddress.Row, rAddress.Column, rAddress.Row ),
    mnHorAlign( nHorAlign )
{
}

bool WorksheetHelperImpl::MergedRange::tryExpand( const CellAddress& rAddress, sal_Int32 nHorAlign )
{
    if( (mnHorAlign == nHorAlign) && (maRange.StartRow == rAddress.Row) &&
        (maRange.EndRow == rAddress.Row) && (maRange.EndColumn + 1 == rAddress.Column) )
    {
        ++maRange.EndColumn;
        return true;
    }
    return false;
}

void WorksheetHelperImpl::writeXfIdRangeProperties( const XfIdRange& rXfIdRange ) const
{
    StylesBuffer& rStyles = getStyles();
    PropertySet aPropSet( getCellRange( rXfIdRange.maRange ) );
    if( rXfIdRange.mnXfId >= 0 )
        rStyles.writeCellXfToPropertySet( aPropSet, rXfIdRange.mnXfId );
    if( rXfIdRange.mnNumFmtId >= 0 )
        rStyles.writeNumFmtToPropertySet( aPropSet, rXfIdRange.mnNumFmtId );
}

void WorksheetHelperImpl::mergeXfIdRanges()
{
    if( !maXfIdRanges.empty() )
    {
        // get row index of last range
        sal_Int32 nLastRow = maXfIdRanges.rbegin()->second.maRange.StartRow;
        // process all ranges located in the same row of the last range
        XfIdRangeMap::iterator aMergeIt = maXfIdRanges.end();
        while( (aMergeIt != maXfIdRanges.begin()) && ((--aMergeIt)->second.maRange.StartRow == nLastRow) )
        {
            const XfIdRange& rMergeXfIdRange = aMergeIt->second;
            // try to find a range that can be merged with rMergeRange
            bool bFound = false;
            for( XfIdRangeMap::iterator aIt = maXfIdRanges.begin(); !bFound && (aIt != aMergeIt); ++aIt )
                if( (bFound = aIt->second.tryMerge( rMergeXfIdRange )) == true )
                    maXfIdRanges.erase( aMergeIt++ );
        }
    }
}

void WorksheetHelperImpl::finalizeXfIdRanges()
{
    // try to merge remaining inserted ranges
    mergeXfIdRanges();
    // write all formatting
    for( XfIdRangeMap::const_iterator aIt = maXfIdRanges.begin(), aEnd = maXfIdRanges.end(); aIt != aEnd; ++aIt )
        writeXfIdRangeProperties( aIt->second );
}

void WorksheetHelperImpl::finalizeHyperlinkRanges() const
{
    for( OoxHyperlinkList::const_iterator aIt = maHyperlinks.begin(), aEnd = maHyperlinks.end(); aIt != aEnd; ++aIt )
    {
        OUStringBuffer aUrlBuffer( aIt->maTarget );
        if( aIt->maLocation.getLength() > 0 )
            aUrlBuffer.append( sal_Unicode( '#' ) ).append( aIt->maLocation );
        OUString aUrl = aUrlBuffer.makeStringAndClear();
        if( aUrl.getLength() > 0 )
        {
            // convert '#SheetName!A1' to '#SheetName.A1'
            if( aUrl[ 0 ] == '#' )
            {
                sal_Int32 nSepPos = aUrl.lastIndexOf( '!' );
                if( nSepPos > 0 )
                {
                    // replace the exclamation mark with a period
                    aUrl = aUrl.replaceAt( nSepPos, 1, OUString( sal_Unicode( '.' ) ) );
                    // #i66592# convert renamed sheets
                    bool bQuotedName = (nSepPos > 3) && (aUrl[ 1 ] == '\'') && (aUrl[ nSepPos - 1 ] == '\'');
                    sal_Int32 nNamePos = bQuotedName ? 2 : 1;
                    sal_Int32 nNameLen = nSepPos - (bQuotedName ? 3 : 1);
                    OUString aSheetName = aUrl.copy( nNamePos, nNameLen );
                    OUString aFinalName = getWorksheets().getFinalSheetName( aSheetName );
                    if( aFinalName.getLength() > 0 )
                        aUrl = aUrl.replaceAt( nNamePos, nNameLen, aFinalName );
                }
            }

            // try to insert URL into each cell of the range
            for( CellAddress aAddress( mnSheet, aIt->maRange.StartColumn, aIt->maRange.StartRow ); aAddress.Row <= aIt->maRange.EndRow; ++aAddress.Row )
                for( aAddress.Column = aIt->maRange.StartColumn; aAddress.Column <= aIt->maRange.EndColumn; ++aAddress.Column )
                    finalizeHyperlink( aAddress, aUrl );
        }
    }
}

void WorksheetHelperImpl::finalizeHyperlink( const CellAddress& rAddress, const OUString& rUrl ) const
{
    Reference< XMultiServiceFactory > xFactory( getDocument(), UNO_QUERY );
    Reference< XCell > xCell = getCell( rAddress );
    Reference< XText > xText( xCell, UNO_QUERY );
    // hyperlinks only supported in text cells
    if( xFactory.is() && xCell.is() && (xCell->getType() == ::com::sun::star::table::CellContentType_TEXT) && xText.is() )
    {
        // create a URL field object and set its properties
        Reference< XTextContent > xUrlField( xFactory->createInstance( maUrlTextField ), UNO_QUERY );
        OSL_ENSURE( xUrlField.is(), "WorksheetHelperImpl::finalizeHyperlink - cannot create text field" );
        if( xUrlField.is() )
        {
            // properties of the URL field
            PropertySet aPropSet( xUrlField );
            aPropSet.setProperty( maUrlProp, rUrl );
            aPropSet.setProperty( maReprProp, xText->getString() );
            try
            {
                // insert the field into the cell
                xText->setString( OUString() );
                Reference< XTextRange > xRange( xText->createTextCursor(), UNO_QUERY_THROW );
                xText->insertTextContent( xRange, xUrlField, sal_False );
            }
            catch( const Exception& )
            {
                OSL_ENSURE( false, "WorksheetHelperImpl::finalizeHyperlink - cannot insert text field" );
            }
        }
    }
}

void WorksheetHelperImpl::finalizeMergedRanges() const
{
    MergedRangeList::const_iterator aIt, aEnd;
    for( aIt = maMergedRanges.begin(), aEnd = maMergedRanges.end(); aIt != aEnd; ++aIt )
        finalizeMergedRange( aIt->maRange );
    for( aIt = maCenterFillRanges.begin(), aEnd = maCenterFillRanges.end(); aIt != aEnd; ++aIt )
        finalizeMergedRange( aIt->maRange );
}

void WorksheetHelperImpl::finalizeMergedRange( const CellRangeAddress& rRange ) const
{
    bool bMultiCol = rRange.StartColumn < rRange.EndColumn;
    bool bMultiRow = rRange.StartRow < rRange.EndRow;

    if( bMultiCol || bMultiRow ) try
    {
        // merge the cell range
        Reference< XMergeable > xMerge( getCellRange( rRange ), UNO_QUERY_THROW );
        xMerge->merge( sal_True );

        // if merging this range worked (no overlapping merged ranges), update cell borders
        PropertySet aTopLeft( getCell( CellAddress( mnSheet, rRange.StartColumn, rRange.StartRow ) ) );

        // copy right border of top-right cell to right border of top-left cell
        if( bMultiCol )
        {
            PropertySet aTopRight( getCell( CellAddress( mnSheet, rRange.EndColumn, rRange.StartRow ) ) );
            BorderLine aLine;
            if( aTopRight.getProperty( aLine, maRightBorderProp ) )
                aTopLeft.setProperty( maRightBorderProp, aLine );
        }

        // copy bottom border of bottom-left cell to bottom border of top-left cell
        if( bMultiRow )
        {
            PropertySet aBottomLeft( getCell( CellAddress( mnSheet, rRange.StartColumn, rRange.EndRow ) ) );
            BorderLine aLine;
            if( aBottomLeft.getProperty( aLine, maBottomBorderProp ) )
                aTopLeft.setProperty( maBottomBorderProp, aLine );
        }
    }
    catch( Exception& )
    {
    }
}

void WorksheetHelperImpl::convertPageStyle()
{
    try
    {
        Reference< XStyleFamiliesSupplier > xFamiliesSup( getDocument(), UNO_QUERY_THROW );
        Reference< XNameAccess > xFamiliesNA( xFamiliesSup->getStyleFamilies(), UNO_QUERY_THROW );
        Reference< XNameContainer > xStylesNC( xFamiliesNA->getByName( CREATE_OUSTRING( "PageStyles" ) ), UNO_QUERY_THROW );

        Reference< XMultiServiceFactory > xFactory( getDocument(), UNO_QUERY_THROW );
        Reference< XStyle > xStyle( xFactory->createInstance( CREATE_OUSTRING( "com.sun.star.style.PageStyle" ) ), UNO_QUERY_THROW );

        Reference< XNamed > xSheetName( mxSheet, UNO_QUERY_THROW );
        OUString aStyleName = CREATE_OUSTRING( "PageStyle_" ) + xSheetName->getName();
        aStyleName = ContainerHelper::insertByUnusedName( xStylesNC, Any( xStyle ), aStyleName, ' ', false );

        // Populate the new page style.
        PropertySet aStyleProps( xStyle );
        maPageStyle.writeToPropertySet( aStyleProps, meSheetType );

        // Link this page style with the current sheet.
        PropertySet aSheetProps( mxSheet );
        aSheetProps.setProperty( CREATE_OUSTRING( "PageStyle" ), aStyleName );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "WorksheetHelperImpl::convertPageStyle - cannot create style sheet" );
    }
}

void WorksheetHelperImpl::convertColumns()
{
    sal_Int32 nNextCol = 0;
    sal_Int32 nMaxCol = mrMaxApiPos.Column;
    // stores first grouped column index for each level
    OutlineLevelVec aColLevels;

    for( OoxColumnDataMap::const_iterator aIt = maColDatas.begin(), aEnd = maColDatas.end(); aIt != aEnd; ++aIt )
    {
        // convert 1-based OOX column indexes to 0-based API column indexes
        sal_Int32 nFirstCol = ::std::max( aIt->second.mnFirstCol - 1, nNextCol );
        sal_Int32 nLastCol = ::std::min( aIt->second.mnLastCol - 1, nMaxCol );

        // process gap between two column datas, use default column data
        if( nNextCol < nFirstCol )
            convertColumns( aColLevels, nNextCol, nFirstCol - 1, maDefColData );
        // process the column data
        convertColumns( aColLevels, nFirstCol, nLastCol, aIt->second );

        // cache next column to be processed
        nNextCol = nLastCol + 1;
    }

    // remaining default columns to end of sheet
    convertColumns( aColLevels, nNextCol, nMaxCol, maDefColData );
    // close remaining column outlines spanning to end of sheet
    convertOutlines( aColLevels, nMaxCol + 1, 0, false, false );
}

void WorksheetHelperImpl::convertColumns( OutlineLevelVec& orColLevels,
        sal_Int32 nFirstCol, sal_Int32 nLastCol, const OoxColumnData& rData )
{
    Reference< XTableColumns > xColumns = getColumns( nFirstCol, nLastCol );
    if( xColumns.is() )
    {
        // set column properties
        PropertySet aPropSet( xColumns );
        // convert 'number of characters' to column width in 1/100 mm
        sal_Int32 nWidth = getUnitConverter().calcMm100FromDigits( rData.mfWidth );
        if( nWidth > 0 )
            aPropSet.setProperty( maWidthProp, nWidth );
        if( rData.mbHidden )
            aPropSet.setProperty( maVisibleProp, false );
    }
    // outline settings for this column range
    convertOutlines( orColLevels, nFirstCol, rData.mnLevel, rData.mbCollapsed, false );
}

void WorksheetHelperImpl::convertRows()
{
    sal_Int32 nNextRow = 0;
    sal_Int32 nMaxRow = mrMaxApiPos.Row;
    // stores first grouped row index for each level
    OutlineLevelVec aRowLevels;

    for( OoxRowDataMap::const_iterator aIt = maRowDatas.begin(), aEnd = maRowDatas.end(); aIt != aEnd; ++aIt )
    {
        // convert 1-based OOX row indexes to 0-based API row indexes
        sal_Int32 nFirstRow = ::std::max( aIt->second.mnFirstRow - 1, nNextRow );
        sal_Int32 nLastRow = ::std::min( aIt->second.mnLastRow - 1, nMaxRow );

        // process gap between two row datas, use default row data
        if( nNextRow < nFirstRow )
            convertRows( aRowLevels, nNextRow, nFirstRow - 1, maDefRowData );
        // process the row data
        convertRows( aRowLevels, nFirstRow, nLastRow, aIt->second );

        // cache next row to be processed
        nNextRow = nLastRow + 1;
    }

    // remaining default rows to end of sheet
    convertRows( aRowLevels, nNextRow, nMaxRow, maDefRowData );
    // close remaining row outlines spanning to end of sheet
    convertOutlines( aRowLevels, nMaxRow + 1, 0, false, true );
}

void WorksheetHelperImpl::convertRows( OutlineLevelVec& orRowLevels,
        sal_Int32 nFirstRow, sal_Int32 nLastRow, const OoxRowData& rData )
{
    Reference< XTableRows > xRows = getRows( nFirstRow, nLastRow );
    if( xRows.is() )
    {
        // set row properties
        PropertySet aPropSet( xRows );
        // convert points to row height in 1/100 mm
        sal_Int32 nHeight = getUnitConverter().calcMm100FromPoints( rData.mfHeight );
        if( nHeight > 0 )
            aPropSet.setProperty( maHeightProp, nHeight );
        if( rData.mbHidden )
            aPropSet.setProperty( maVisibleProp, false );
    }
    // outline settings for this row range
    convertOutlines( orRowLevels, nFirstRow, rData.mnLevel, rData.mbCollapsed, true );
}

void WorksheetHelperImpl::convertOutlines( OutlineLevelVec& orLevels,
        sal_Int32 nColRow, sal_Int32 nLevel, bool bCollapsed, bool bRows )
{
    /*  It is ensured from caller functions, that this function is called
        without any gaps between the processed column or row ranges. */

    OSL_ENSURE( nLevel >= 0, "WorksheetHelperImpl::convertOutlines - negative outline level" );
    nLevel = ::std::max< sal_Int32 >( nLevel, 0 );

    sal_Int32 nSize = orLevels.size();
    if( nSize < nLevel )
    {
        // Outline level increased. Push the begin column position.
        for( sal_Int32 nIndex = nSize; nIndex < nLevel; ++nIndex )
            orLevels.push_back( nColRow );
    }
    else if( nLevel < nSize )
    {
        // Outline level decreased. Pop them all out.
        for( sal_Int32 nIndex = nLevel; nIndex < nSize; ++nIndex )
        {
            sal_Int32 nFirstInLevel = orLevels.back();
            orLevels.pop_back();
            groupColumnsOrRows( nFirstInLevel, nColRow - 1, bCollapsed, bRows );
            bCollapsed = false; // collapse only once
        }
    }
}

void WorksheetHelperImpl::groupColumnsOrRows( sal_Int32 nFirstColRow, sal_Int32 nLastColRow, bool bCollapse, bool bRows )
{
    Reference< XSheetOutline > xOutline( mxSheet, UNO_QUERY );
    if( xOutline.is() )
    {
        if( bRows )
        {
            CellRangeAddress aRange( mnSheet, 0, nFirstColRow, 0, nLastColRow );
            xOutline->group( aRange, ::com::sun::star::table::TableOrientation_ROWS );
            if( bCollapse )
                xOutline->hideDetail( aRange );
        }
        else
        {
            CellRangeAddress aRange( mnSheet, nFirstColRow, 0, nLastColRow, 0 );
            xOutline->group( aRange, ::com::sun::star::table::TableOrientation_COLUMNS );
            if( bCollapse )
                xOutline->hideDetail( aRange );
        }
    }
}

// ============================================================================
// ============================================================================

WorksheetHelper::WorksheetHelper( const GlobalDataHelper& rGlobalData,
        WorksheetType eSheetType, const Reference< XSpreadsheet >& rxSheet, sal_Int32 nSheet ) :
    GlobalDataHelper( rGlobalData ),
    mxImpl( new WorksheetHelperImpl( rGlobalData, eSheetType, rxSheet, nSheet ) )
{
}

WorksheetHelper::~WorksheetHelper()
{
}

WorksheetType WorksheetHelper::getSheetType() const
{
    return mxImpl->getSheetType();
}

Reference< XSpreadsheet > WorksheetHelper::getXSpreadsheet() const
{
    return mxImpl->getXSpreadsheet();
}

sal_Int16 WorksheetHelper::getSheetIndex() const
{
    return mxImpl->getSheetIndex();
}

sal_Int32 WorksheetHelper::incCondFormatIndex() const
{
    return mxImpl->incCondFormatIndex();
}

sal_Int32 WorksheetHelper::getCondFormatIndex() const
{
    return mxImpl->getCondFormatIndex();
}

Reference< XCell > WorksheetHelper::getCell( const CellAddress& rAddress ) const
{
    return mxImpl->getCell( rAddress );
}

Reference< XCell > WorksheetHelper::getCell( const OUString& rAddressStr, CellAddress* opAddress ) const
{
    CellAddress aAddress;
    if( getAddressConverter().convertToCellAddress( aAddress, rAddressStr, mxImpl->getSheetIndex(), true ) )
    {
        if( opAddress ) *opAddress = aAddress;
        return mxImpl->getCell( aAddress );
    }
    return Reference< XCell >();
}

Reference< XCell > WorksheetHelper::getCell( const BiffAddress& rBiffAddress, CellAddress* opAddress ) const
{
    CellAddress aAddress;
    if( getAddressConverter().convertToCellAddress( aAddress, rBiffAddress, mxImpl->getSheetIndex(), true ) )
    {
        if( opAddress ) *opAddress = aAddress;
        return mxImpl->getCell( aAddress );
    }
    return Reference< XCell >();
}

Reference< XCellRange > WorksheetHelper::getCellRange( const CellRangeAddress& rRange ) const
{
    return mxImpl->getCellRange( rRange );
}

Reference< XCellRange > WorksheetHelper::getCellRange( const OUString& rRangeStr, CellRangeAddress* opRange ) const
{
    CellRangeAddress aRange;
    if( getAddressConverter().convertToCellRange( aRange, rRangeStr, mxImpl->getSheetIndex(), true ) )
    {
        if( opRange ) *opRange = aRange;
        return mxImpl->getCellRange( aRange );
    }
    return Reference< XCellRange >();
}

Reference< XCellRange > WorksheetHelper::getCellRange( const BiffRange& rBiffRange, CellRangeAddress* opRange ) const
{
    CellRangeAddress aRange;
    if( getAddressConverter().convertToCellRange( aRange, rBiffRange, mxImpl->getSheetIndex(), true ) )
    {
        if( opRange ) *opRange = aRange;
        return mxImpl->getCellRange( aRange );
    }
    return Reference< XCellRange >();
}

CellAddress WorksheetHelper::getCellAddress( const Reference< XCell >& rxCell )
{
    CellAddress aAddress;
    Reference< XCellAddressable > xAddressable( rxCell, UNO_QUERY );
    OSL_ENSURE( xAddressable.is(), "WorksheetHelper::getCellAddress - cell reference not addressable" );
    if( xAddressable.is() )
        aAddress = xAddressable->getCellAddress();
    return aAddress;
}

CellRangeAddress WorksheetHelper::getRangeAddress( const Reference< XCellRange >& rxRange )
{
    CellRangeAddress aRange;
    Reference< XCellRangeAddressable > xAddressable( rxRange, UNO_QUERY );
    OSL_ENSURE( xAddressable.is(), "WorksheetHelper::getRangeAddress - cell range reference not addressable" );
    if( xAddressable.is() )
        aRange = xAddressable->getRangeAddress();
    return aRange;
}

Reference< XCellRange > WorksheetHelper::getColumn( sal_Int32 nCol ) const
{
    return mxImpl->getColumn( nCol );
}

Reference< XCellRange > WorksheetHelper::getRow( sal_Int32 nRow ) const
{
    return mxImpl->getRow( nRow );
}

Reference< XTableColumns > WorksheetHelper::getColumns( sal_Int32 nFirstCol, sal_Int32 nLastCol ) const
{
    return mxImpl->getColumns( nFirstCol, nLastCol );
}

Reference< XTableRows > WorksheetHelper::getRows( sal_Int32 nFirstRow, sal_Int32 nLastRow ) const
{
    return mxImpl->getRows( nFirstRow, nLastRow );
}

SharedFormulaBuffer& WorksheetHelper::getSharedFormulas() const
{
    return mxImpl->getSharedFormulas();
}

PageStyle& WorksheetHelper::getPageStyle() const
{
    return mxImpl->getPageStyle();
}

PhoneticSettings& WorksheetHelper::getPhoneticSettings() const
{
    return mxImpl->getPhoneticSettings();
}

OoxSheetViewData& WorksheetHelper::createSheetViewData() const
{
    return getViewSettings().createSheetViewData( mxImpl->getSheetIndex() );
}

void WorksheetHelper::setBooleanCell( const Reference< XCell >& rxCell, bool bValue ) const
{
    OSL_ENSURE( rxCell.is(), "WorksheetHelper::setBooleanCell - missing cell interface" );
    rxCell->setFormula( mxImpl->getBooleanFormula( bValue ) );
}

void WorksheetHelper::setErrorCell( const Reference< XCell >& rxCell, const OUString& rErrorCode ) const
{
    setErrorCell( rxCell, mxImpl->getBiffErrorCode( rErrorCode ) );
}

void WorksheetHelper::setErrorCell( const Reference< XCell >& rxCell, sal_uInt8 nErrorCode ) const
{
    Reference< XFormulaTokens > xTokens( rxCell, UNO_QUERY );
    OSL_ENSURE( xTokens.is(), "WorksheetHelper::setErrorCell - missing cell interface" );
    if( xTokens.is() )
    {
        SimpleFormulaContext aContext( xTokens );
        aContext.setBaseAddress( getCellAddress( rxCell ) );
        getFormulaParser().convertErrorToFormula( aContext, nErrorCode );
    }
}

void WorksheetHelper::setCell( OoxCellData& orCellData ) const
{
    OSL_ENSURE( orCellData.mxCell.is(), "WorksheetHelper::setCell - missing cell interface" );
    if( orCellData.mbHasValueStr ) switch( orCellData.mnCellType )
    {
        case XML_b:
            setBooleanCell( orCellData.mxCell, orCellData.maValueStr.toDouble() != 0.0 );
            // #108770# set 'Standard' number format for all Boolean cells
            orCellData.mnNumFmtId = 0;
        break;

        case XML_n:
            orCellData.mxCell->setValue( orCellData.maValueStr.toDouble() );
        break;

        case XML_e:
            setErrorCell( orCellData.mxCell, orCellData.maValueStr );
        break;

        case XML_str:
        {
            Reference< XText > xText( orCellData.mxCell, UNO_QUERY );
            if( xText.is() )
                xText->setString( orCellData.maValueStr );
        }
        break;

        case XML_s:
        {
            Reference< XText > xText( orCellData.mxCell, UNO_QUERY );
            if( xText.is() )
                getSharedStrings().convertString( xText, orCellData.maValueStr.toInt32(), orCellData.mnXfId );
        }
        break;
    }
}

void WorksheetHelper::setCellFormat( const OoxCellData& rCellData )
{
    mxImpl->setCellFormat( rCellData );
}

void WorksheetHelper::setHyperlink( const OoxHyperlinkData& rHyperlink )
{
    mxImpl->setHyperlink( rHyperlink );
}

void WorksheetHelper::setMergedRange( const CellRangeAddress& rRange )
{
    mxImpl->setMergedRange( rRange );
}

void WorksheetHelper::setBaseColumnWidth( sal_Int32 nWidth )
{
    mxImpl->setBaseColumnWidth( nWidth );
}

void WorksheetHelper::setDefaultColumnWidth( double fWidth )
{
    mxImpl->setDefaultColumnWidth( fWidth );
}

void WorksheetHelper::setColumnData( const OoxColumnData& rData )
{
    mxImpl->setColumnData( rData );
}

void WorksheetHelper::setDefaultRowSettings( double fHeight, bool bHidden )
{
    mxImpl->setDefaultRowSettings( fHeight, bHidden );
}

void WorksheetHelper::setRowData( const OoxRowData& rData )
{
    mxImpl->setRowData( rData );
}

void WorksheetHelper::setOutlineSummarySymbols( bool bSummaryRight, bool bSummaryBelow )
{
    mxImpl->setOutlineSummarySymbols( bSummaryRight, bSummaryBelow );
}

void WorksheetHelper::convertColumnFormat( sal_Int32 nFirstCol, sal_Int32 nLastCol, sal_Int32 nXfId )
{
    mxImpl->convertColumnFormat( nFirstCol, nLastCol, nXfId );
}

void WorksheetHelper::convertRowFormat( sal_Int32 nFirstRow, sal_Int32 nLastRow, sal_Int32 nXfId )
{
    mxImpl->convertRowFormat( nFirstRow, nLastRow, nXfId );
}

void WorksheetHelper::convertPageBreak( const OoxPageBreakData& rData, bool bRowBreak )
{
    mxImpl->convertPageBreak( rData, bRowBreak );
}

void WorksheetHelper::finalizeWorksheetImport()
{
    mxImpl->finalizeWorksheetImport();
}

// ============================================================================

} // namespace xls
} // namespace oox

