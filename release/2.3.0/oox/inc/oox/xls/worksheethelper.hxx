/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: worksheethelper.hxx,v $
 *
 *  $Revision: 1.1.2.32 $
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

#ifndef OOX_XLS_WORKSHEETHELPER_HXX
#define OOX_XLS_WORKSHEETHELPER_HXX

#include <com/sun/star/table/CellAddress.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include "oox/xls/globaldatahelper.hxx"

namespace com { namespace sun { namespace star {
    namespace table { class XTableColumns; }
    namespace table { class XTableRows; }
    namespace table { class XCell; }
    namespace table { class XCellRange; }
    namespace sheet { class XSpreadsheet; }
    namespace sheet { class XSheetOutline; }
    namespace sheet { class XNamedRange; }
} } }

namespace oox {
namespace xls {

struct BiffAddress;
struct BiffRange;
class PageStyle;
class PhoneticSettings;
struct OoxSheetViewData;

// ============================================================================

/** An enumeration for all types of sheets in a workbook. */
enum WorksheetType
{
    SHEETTYPE_WORKSHEET,                    /// Worksheet.
    SHEETTYPE_CHART,                        /// Chart sheet.
    SHEETTYPE_MACRO                         /// BIFF4 macro sheet.
};

// ============================================================================

/** Stores some data about a cell. */
struct OoxCellData
{
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell > mxCell;
    ::com::sun::star::table::CellAddress maAddress;
    ::rtl::OUString     maValueStr;         /// String containing cell value data.
    ::rtl::OUString     maFormulaStr;       /// String containing formula definition.
    ::rtl::OUString     maFormulaRef;       /// String containing formula range for array/shared formulas.
    sal_Int32           mnCellType;         /// Data type of the cell.
    sal_Int32           mnFormulaType;      /// Type of the imported formula.
    sal_Int32           mnSharedId;         /// Shared formula identifier for current cell.
    sal_Int32           mnXfId;             /// XF identifier for the cell.
    sal_Int32           mnNumFmtId;         /// Forced number format for the cell.
    bool                mbHasValueStr;      /// True = contents of maValueStr are valid.

    inline explicit     OoxCellData() { reset(); }
    void                reset();
};

// ----------------------------------------------------------------------------

/** Stores formatting data about a range of columns. */
struct OoxColumnData
{
    sal_Int32           mnFirstCol;         /// 1-based (!) index of first column.
    sal_Int32           mnLastCol;          /// 1-based (!) index of last column.
    double              mfWidth;            /// Column width in number of characters.
    sal_Int32           mnXfId;             /// Column default formatting.
    sal_Int32           mnLevel;            /// Column outline level.
    bool                mbHidden;           /// True = column is hidden.
    bool                mbCollapsed;        /// True = column outline is collapsed.

    explicit            OoxColumnData();

    /** Expands this entry with the passed column range, if column settings are equal. */
    bool                tryExpand( const OoxColumnData& rNewData );
};

// ----------------------------------------------------------------------------

/** Stores formatting data about a range of rows. */
struct OoxRowData
{
    sal_Int32           mnFirstRow;         /// 1-based (!) index of first row.
    sal_Int32           mnLastRow;          /// 1-based (!) index of last row.
    double              mfHeight;           /// Row height in points.
    sal_Int32           mnXfId;             /// Row default formatting (see mbIsFormatted).
    sal_Int32           mnLevel;            /// Row outline level.
    bool                mbFormatted;        /// Cells in row have default formatting.
    bool                mbHidden;           /// True = row is hidden.
    bool                mbCollapsed;        /// True = row outline is collapsed.

    explicit            OoxRowData();

    /** Expands this entry with the passed row range, if row settings are equal. */
    bool                tryExpand( const OoxRowData& rNewData );
};

// ----------------------------------------------------------------------------

/** Stores formatting data about a page break. */
struct OoxPageBreakData
{
    sal_Int32           mnColRow;           /// 0-based (!) index of column/row.
    bool                mbManual;           /// True = manual page break.

    explicit            OoxPageBreakData();
};

// ----------------------------------------------------------------------------

/** Stores data about a hyperlink range. */
struct OoxHyperlinkData
{
    ::com::sun::star::table::CellRangeAddress maRange;
    ::rtl::OUString     maTarget;
    ::rtl::OUString     maLocation;
    ::rtl::OUString     maDisplay;
    ::rtl::OUString     maTooltip;

    explicit            OoxHyperlinkData();
};

// ============================================================================
// ============================================================================

class SharedFormulaBuffer : public GlobalDataHelper
{
public:
    explicit            SharedFormulaBuffer( const GlobalDataHelper& rGlobalData );

    /** Imports a shared formula from a OOX formula string. */
    void                importSharedFmla( const ::rtl::OUString& rFormula, sal_Int32 nId,
                            const ::com::sun::star::table::CellAddress& rBaseAddr );
    /** Imports a shared formula from a SHRFMLA record in the passed stream. */
    void                importShrFmla( BiffInputStream& rStrm,
                            const ::com::sun::star::table::CellAddress& rBaseAddr );

    /** Returns the token index of the defined name representing the specified shared formula. */
    sal_Int32           getTokenIndexFromId( sal_Int32 nSharedId ) const;

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XNamedRange >
                        createDefinedName( sal_Int32 nSheet, sal_Int32 nSharedId );

private:
    typedef ::std::map< sal_Int32, sal_Int32 > TokenIndexMap;
    TokenIndexMap       maIndexMap;         /// Maps shared formula identifier to defined name identifier.
};

// ============================================================================

class WorksheetHelperImpl;

class WorksheetHelper : public GlobalDataHelper
{
public:
    explicit            WorksheetHelper(
                            const GlobalDataHelper& rGlobalData,
                            WorksheetType eSheetType,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >& rxSheet,
                            sal_Int32 nSheet );

    virtual             ~WorksheetHelper();

    /** Return this helper for better code readability in derived classes. */
    inline const WorksheetHelper& getWorksheetHelper() const { return *this; }
    /** Return this helper for better code readability in derived classes. */
    inline WorksheetHelper& getWorksheetHelper() { return *this; }

    /** Returns the type of this sheet. */
    WorksheetType       getSheetType() const;
    /** Returns the XSpreadsheet interface of the processed sheet. */
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >
                        getXSpreadsheet() const;
    /** Returns the index of the current sheet. */
    sal_Int16           getSheetIndex() const;

    /** Increment the conditional formatting index by one, and returns the old
        index. */
    sal_Int32           incCondFormatIndex() const;
    /** Returns the index of the current conditional formatting item. */
    sal_Int32           getCondFormatIndex() const;

    /** Returns the XCell interface for the passed cell address. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >
                        getCell(
                            const ::com::sun::star::table::CellAddress& rAddress ) const;
    /** Returns the XCell interface for the passed cell address string. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >
                        getCell(
                            const ::rtl::OUString& rAddressStr,
                            ::com::sun::star::table::CellAddress* opAddress = 0 ) const;
    /** Returns the XCell interface for the passed BIFF cell address. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >
                        getCell(
                            const BiffAddress& rBiffAddress,
                            ::com::sun::star::table::CellAddress* opAddress = 0 ) const;

    /** Returns the XCellRange interface for the passed cell range address. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >
                        getCellRange(
                            const ::com::sun::star::table::CellRangeAddress& rRange ) const;
    /** Returns the XCellRange interface for the passed range address string. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >
                        getCellRange(
                            const ::rtl::OUString& rRangeStr,
                            ::com::sun::star::table::CellRangeAddress* opRange = 0 ) const;
    /** Returns the XCellRange interface for the passed BIFF range address. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >
                        getCellRange(
                            const BiffRange& rBiffRange,
                            ::com::sun::star::table::CellRangeAddress* opRange = 0 ) const;

    /** Returns the address of the passed cell. The cell reference must be valid. */
    static ::com::sun::star::table::CellAddress
                        getCellAddress(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >& rxCell );
    /** Returns the address of the passed cell range. The range reference must be valid. */
    static ::com::sun::star::table::CellRangeAddress
                        getRangeAddress(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >& rxRange );

    /** Returns the XCellRange interface for a column. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >
                        getColumn( sal_Int32 nCol ) const;
    /** Returns the XCellRange interface for a row. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XCellRange >
                        getRow( sal_Int32 nRow ) const;

    /** Returns the XTableColumns interface for a range of columns. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XTableColumns >
                        getColumns( sal_Int32 nFirstCol, sal_Int32 nLastCol ) const;
    /** Returns the XTableRows interface for a range of rows. */
    ::com::sun::star::uno::Reference< ::com::sun::star::table::XTableRows >
                        getRows( sal_Int32 nFirstRow, sal_Int32 nLastRow ) const;

    /** Returns the buffer containing all shared formulas in this sheet. */
    SharedFormulaBuffer& getSharedFormulas() const;
    /** Returns the page style object that implements import/export of page/print settings. */
    PageStyle&          getPageStyle() const;
    /** Returns the global phonetic settings for this sheet. */
    PhoneticSettings&   getPhoneticSettings() const;
    /** Returns the structure containing view settings for this sheet. */
    OoxSheetViewData&   createSheetViewData() const;

    /** Sets the passed boolean value to the cell and adjusts number format. */
    void                setBooleanCell(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >& rxCell,
                            bool bValue ) const;
    /** Sets the passed BIFF error code to the cell (by converting it to a formula). */
    void                setErrorCell(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >& rxCell,
                            const ::rtl::OUString& rErrorCode ) const;
    /** Sets the passed BIFF error code to the cell (by converting it to a formula). */
    void                setErrorCell(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::table::XCell >& rxCell,
                            sal_uInt8 nErrorCode ) const;
    /** Sets cell contents to the cell specified in the passed cell data object. */
    void                setCell( OoxCellData& orCellData ) const;

    /** Stores the cell formatting data of the current cell. */
    void                setCellFormat( const OoxCellData& rCellData );
    /** Inserts the hyperlink URL into the passed cell range. */
    void                setHyperlink( const OoxHyperlinkData& rHyperlink );
    /** Merges the cells in the passed cell range. */
    void                setMergedRange( const ::com::sun::star::table::CellRangeAddress& rRange );

    /** Sets base width for all columns (without padding pixels). This value
        is only used, if width has not been set with setDefaultColumnWidth(). */
    void                setBaseColumnWidth( sal_Int32 nWidth );
    /** Sets default width for all columns. This function overrides the base
        width set with the setBaseColumnWidth() function. */
    void                setDefaultColumnWidth( double fWidth );
    /** Sets column settings for a specific range of columns.
        @descr  Column default formatting is converted directly, other settings
        are cached and converted in the finalizeWorksheetImport() call. */
    void                setColumnData( const OoxColumnData& rData );

    /** Sets default height and hidden state for all unused rows in the sheet. */
    void                setDefaultRowSettings( double fHeight, bool bHidden );
    /** Sets row settings for a specific range of rows.
        @descr  Row default formatting is converted directly, other settings
        are cached and converted in the finalizeWorksheetImport() call. */
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
    ::boost::shared_ptr< WorksheetHelperImpl > mxImpl;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

