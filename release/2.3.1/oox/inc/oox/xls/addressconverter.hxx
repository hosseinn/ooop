/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: addressconverter.hxx,v $
 *
 *  $Revision: 1.1.2.12 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/29 14:01:26 $
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

#ifndef OOX_XLS_ADDRESSCONVERTER_HXX
#define OOX_XLS_ADDRESSCONVERTER_HXX

#include <vector>
#include <com/sun/star/table/CellAddress.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include "oox/xls/globaldatahelper.hxx"

namespace oox {
namespace xls {

// ============================================================================

/** A 2D cell address struct with Excel column and row indexes. */
struct BiffAddress
{
    sal_uInt16          mnCol;
    sal_uInt16          mnRow;

    inline explicit     BiffAddress() : mnCol( 0 ), mnRow( 0 ) {}
    inline explicit     BiffAddress( sal_uInt16 nCol, sal_uInt16 nRow ) : mnCol( nCol ), mnRow( nRow ) {}

    inline void         set( sal_uInt16 nCol, sal_uInt16 nRow ) { mnCol = nCol; mnRow = nRow; }

    void                read( BiffInputStream& rStrm, bool bCol16Bit = true );
    void                write( BiffOutputStream& rStrm, bool bCol16Bit = true ) const;
};

bool operator==( const BiffAddress& rL, const BiffAddress& rR );
bool operator<( const BiffAddress& rL, const BiffAddress& rR );
BiffInputStream& operator>>( BiffInputStream& rStrm, BiffAddress& rPos );
BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffAddress& rPos );

// ----------------------------------------------------------------------------

/** A 2D cell range address struct with Excel column and row indexes. */
struct BiffRange
{
    BiffAddress         maFirst;
    BiffAddress         maLast;

    inline explicit     BiffRange() {}
    inline explicit     BiffRange( const BiffAddress& rPos ) : maFirst( rPos ), maLast( rPos ) {}
    inline explicit     BiffRange( const BiffAddress& rFirst, const BiffAddress& rLast ) : maFirst( rFirst ), maLast( rLast ) {}
    inline explicit     BiffRange( sal_uInt16 nCol1, sal_uInt16 nRow1, sal_uInt16 nCol2, sal_uInt16 nRow2 ) :
                            maFirst( nCol1, nRow1 ), maLast( nCol2, nRow2 ) {}

    inline void         set( const BiffAddress& rFirst, const BiffAddress& rLast )
                            { maFirst = rFirst; maLast = rLast; }
    inline void         set( sal_uInt16 nCol1, sal_uInt16 nRow1, sal_uInt16 nCol2, sal_uInt16 nRow2 )
                            { maFirst.set( nCol1, nRow1 ); maLast.set( nCol2, nRow2 ); }

    inline sal_uInt16   getColCount() const { return maLast.mnCol - maFirst.mnCol + 1; }
    inline sal_uInt16   getRowCount() const { return maLast.mnRow - maFirst.mnRow + 1; }
    bool                contains( const BiffAddress& rPos ) const;

    void                read( BiffInputStream& rStrm, bool bCol16Bit = true );
    void                write( BiffOutputStream& rStrm, bool bCol16Bit = true ) const;
};

bool operator==( const BiffRange& rL, const BiffRange& rR );
bool operator<( const BiffRange& rL, const BiffRange& rR );
BiffInputStream& operator>>( BiffInputStream& rStrm, BiffRange& rRange );
BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffRange& rRange );

// ----------------------------------------------------------------------------

/** A 2D cell range address list with Excel column and row indexes. */
class BiffRangeList : public ::std::vector< BiffRange >
{
public:
    inline explicit     BiffRangeList() {}

    BiffRange           getEnclosingRange() const;

    void                read( BiffInputStream& rStrm, bool bCol16Bit = true );
    void                write( BiffOutputStream& rStrm, bool bCol16Bit = true ) const;
    void                writeSubList( BiffOutputStream& rStrm,
                            size_t nBegin, size_t nCount, bool bCol16Bit = true ) const;
};

BiffInputStream& operator>>( BiffInputStream& rStrm, BiffRangeList& rRanges );
BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffRangeList& rRanges );

// ============================================================================
// ============================================================================

/** Converter for cell addresses and cell ranges for OOX and BIFF filters.
 */
class AddressConverter : public GlobalDataHelper
{
public:
    explicit            AddressConverter( const GlobalDataHelper& rGlobalData );

    // ------------------------------------------------------------------------

    /** Tries to parse the passed string for a 2d cell address in A1 notation.

        This function accepts all strings that match the regular expression
        "[a-zA-Z]{1,6}0*[1-9][0-9]{0,8}" (without quotes), i.e. 1 to 6 letters
        for the column index (translated to 0-based column indexes from 0 to
        321,272,405), and 1 to 9 digits for the 1-based row index (translated
        to 0-based row indexes from 0 to 999,999,998). The row number part may
        contain leading zeros, they will be ignored. It is up to the caller to
        handle cell addresses outside of a specific valid range (e.g. the
        entire spreadsheet).

        @param ornColumn  (out-parameter) Returns the converted column index.
        @param ornRow  (out-parameter) returns the converted row index.
        @param rString  The string containing the cell address.
        @param nStart  Start index of string part in rString to be parsed.
        @param nLength  Length of string part in rString to be parsed.

        @return  true = Parsed string was valid, returned values can be used.
     */
    static bool         parseAddress2d(
                            sal_Int32& ornColumn, sal_Int32& ornRow,
                            const ::rtl::OUString& rString,
                            sal_Int32 nStart = 0,
                            sal_Int32 nLength = SAL_MAX_INT32 );

    /** Tries to parse the passed string for a 2d cell range in A1 notation.

        This function accepts all strings that match the regular expression
        "ADDR(:ADDR)?" (without quotes), where ADDR is a cell address accepted
        by the parseAddress2d() function of this class. It is up to the caller
        to handle cell ranges outside of a specific valid range (e.g. the
        entire spreadsheet).

        @param ornStartColumn  (out-parameter) Returns the converted start column index.
        @param ornStartRow  (out-parameter) returns the converted start row index.
        @param ornEndColumn  (out-parameter) Returns the converted end column index.
        @param ornEndRow  (out-parameter) returns the converted end row index.
        @param rString  The string containing the cell address.
        @param nStart  Start index of string part in rString to be parsed.
        @param nLength  Length of string part in rString to be parsed.

        @return  true = Parsed string was valid, returned values can be used.
     */
    static bool         parseRange2d(
                            sal_Int32& ornStartColumn, sal_Int32& ornStartRow,
                            sal_Int32& ornEndColumn, sal_Int32& ornEndRow,
                            const ::rtl::OUString& rString,
                            sal_Int32 nStart = 0,
                            sal_Int32 nLength = SAL_MAX_INT32 );

    /** Tries to parse an encoded name of an external link target in BIFF
        documents, e.g. from EXTERNSHEET, SUPBOOK, or DCONREF records.

        @param orClassName  (out-parameter) DDE server name or OLE class name.
        @param orTargetUrl  (out-parameter) Traget URL, DDE topic or OLE object name.
        @param orSheetName  (out-parameter) Sheet name in target document.
        @param rBiffEncoded  Encoded name of the external link target.

        @return  true = Parsed string was valid, returned strings can be used.
      */
    bool                parseEncodedTarget(
                            ::rtl::OUString& orClassName,
                            ::rtl::OUString& orTargetUrl,
                            ::rtl::OUString& orSheetName,
                            const ::rtl::OUString& rBiffEncoded );

    // ------------------------------------------------------------------------

    /** Returns the biggest valid cell address in the own Calc document. */
    inline const ::com::sun::star::table::CellAddress&
                        getMaxApiAddress() const { return maMaxApiPos; }

    /** Returns the biggest valid cell address in the imported/exported
        Excel document. */
    inline const ::com::sun::star::table::CellAddress&
                        getMaxXlsAddress() const { return maMaxXlsPos; }

    /** Returns the biggest valid cell address in both Calc and the
        imported/exported Excel document. */
    inline const ::com::sun::star::table::CellAddress&
                        getMaxAddress() const { return maMaxPos; }

    /** Returns the column overflow status. */
    inline bool         isColOverflow() const { return mbColOverflow; }
    /** Returns the row overflow status. */
    inline bool         isRowOverflow() const { return mbRowOverflow; }
    /** Returns the sheet overflow status. */
    inline bool         isTabOverflow() const { return mbTabOverflow; }

    // ------------------------------------------------------------------------

    /** Checks if the passed column index is valid.

        @param nCol  The column index to check.
        @param bTrackOverflow  true = Update the internal overflow flag, if the
            column index is outside of the supported limits.
        @return  true = Passed column index is valid (no index overflow).
     */
    bool                checkCol( sal_Int32 nCol, bool bTrackOverflow );

    /** Checks if the passed row index is valid.

        @param nRow  The row index to check.
        @param bTrackOverflow  true = Update the internal overflow flag, if the
            row index is outside of the supported limits.
        @return  true = Passed row index is valid (no index overflow).
     */
    bool                checkRow( sal_Int32 nRow, bool bTrackOverflow );

    /** Checks if the passed sheet index is valid.

        @param nSheet  The sheet index to check.
        @param bTrackOverflow  true = Update the internal overflow flag, if the
            sheet index is outside of the supported limits.
        @return  true = Passed sheet index is valid (no index overflow).
     */
    bool                checkTab( sal_Int16 nSheet, bool bTrackOverflow );

    // ------------------------------------------------------------------------

    /** Checks the passed cell address if it fits into the spreadsheet limits.

        @param rAddress  The cell address to be checked.
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the address is outside of the supported sheet limits.
        @return  true = Passed address is valid (no index overflow).
     */
    bool                checkCellAddress(
                            const ::com::sun::star::table::CellAddress& rAddress,
                            bool bTrackOverflow );

    /** Tries to convert the passed string to a single cell address.

        @param orAddress  (out-parameter) Returns the converted cell address.
        @param rString  Cell address string in A1 notation.
        @param nSheet  Sheet index to be inserted into orAddress (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the address is outside of the supported sheet limits.
        @return  true = Converted address is valid (no index overflow).
     */
    bool                convertToCellAddress(
                            ::com::sun::star::table::CellAddress& orAddress,
                            const ::rtl::OUString& rString,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    /** Tries to convert the passed BIFF address to a single cell address.

        @param orAddress  (out-parameter) Returns the converted cell address.
        @param rBiffAddress  BIFF cell address struct.
        @param nSheet  Sheet index to be inserted into orAddress (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the address is outside of the supported sheet limits.
        @return  true = Converted address is valid (no index overflow).
     */
    bool                convertToCellAddress(
                            ::com::sun::star::table::CellAddress& orAddress,
                            const BiffAddress& rBiffAddress,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    /** Returns a valid cell address by moving it into allowed dimensions.

        @param rString  Cell address string in A1 notation.
        @param nSheet  Sheet index for the returned address (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the address is outside of the supported sheet limits.
        @return  A valid API cell address struct. */
    ::com::sun::star::table::CellAddress
                        createValidCellAddress(
                            const ::rtl::OUString& rString,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    /** Returns a valid cell address by moving it into allowed dimensions.

        @param rBiffAddress  BIFF cell address struct.
        @param nSheet  Sheet index for the returned address (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the address is outside of the supported sheet limits.
        @return  A valid API cell address struct. */
    ::com::sun::star::table::CellAddress
                        createValidCellAddress(
                            const BiffAddress& rBiffAddress,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    // ------------------------------------------------------------------------

    /** Checks the passed cell range if it fits into the spreadsheet limits.

        @param rRange  The cell range address to be checked.
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the passed range contains cells outside of the supported sheet
            limits.
        @return  true = Cell range is valid. This function returns also true,
            if only parts of the range are outside the current sheet limits.
            Returns false, if the entire range is outside the sheet limits.
     */
    bool                checkCellRange(
                            const ::com::sun::star::table::CellRangeAddress& rRange,
                            bool bTrackOverflow );

    /** Tries to restrict the passed cell range to current sheet limits.

        @param orRange  (in-out-parameter) Converts the passed cell range
            into a valid cell range address. If the passed range contains cells
            outside the currently supported spreadsheet limits, it will be
            cropped to these limits.
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original range contains cells outside of the supported sheet
            limits.
        @return  true = Converted range address is valid. This function
            returns also true, if the range has been cropped, but still
            contains cells inside the current sheet limits. Returns false, if
            the entire range is outside the sheet limits.
     */
    bool                validateCellRange(
                            ::com::sun::star::table::CellRangeAddress& orRange,
                            bool bTrackOverflow );

    /** Tries to convert the passed string to a cell range address.

        @param orRange  (out-parameter) Returns the converted cell range
            address. If the original range in the passed string contains cells
            outside the currently supported spreadsheet limits, it will be
            cropped to these limits. Example: the range string "A1:ZZ100000"
            may be converted to the range A1:IV65536.
        @param rString  Cell range string in A1 notation.
        @param nSheet  Sheet index to be inserted into orRange (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original range contains cells outside of the supported sheet
            limits.
        @return  true = Converted and returned range is valid. This function
            returns also true, if the range has been cropped, but still
            contains cells inside the current sheet limits. Returns false, if
            the entire range is outside the sheet limits.
     */
    bool                convertToCellRange(
                            ::com::sun::star::table::CellRangeAddress& orRange,
                            const ::rtl::OUString& rString,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    /** Tries to convert the passed BIFF range to a cell range address.

        @param orRange  (out-parameter) Returns the converted cell range
            address. If the original range in the passed string contains cells
            outside the currently supported spreadsheet limits, it will be
            cropped to these limits.
        @param rBiffRange  BIFF cell range struct.
        @param nSheet  Sheet index to be inserted into orRange (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original range contains cells outside of the supported sheet
            limits.
        @return  true = Converted and returned range is valid. This function
            returns also true, if the range has been cropped, but still
            contains cells inside the current sheet limits. Returns false, if
            the entire range is outside the sheet limits.
     */
    bool                convertToCellRange(
                            ::com::sun::star::table::CellRangeAddress& orRange,
                            const BiffRange& rBiffRange,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    // ------------------------------------------------------------------------

    /** Checks the passed cell range list if it fits into the spreadsheet limits.

        @param rRanges  The cell range list to be checked.
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the passed range list contains cells outside of the supported sheet
            limits.
        @return  true = All cell ranges are valid. This function returns also
            true, if only parts of the ranges are outside the current sheet
            limits. Returns false, if one of the ranges is completely outside
            the sheet limits.
     */
    bool                checkCellRangeList(
                            const ::std::vector< ::com::sun::star::table::CellRangeAddress >& rRanges,
                            bool bTrackOverflow );

    /** Tries to restrict the passed cell range list to current sheet limits.

        @param orRanges  (in-out-parameter) Restricts the cell range addresses
            in the passed list to the current sheet limits and removes invalid
            ranges from the list.
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original ranges contain cells outside of the supported sheet
            limits.
     */
    void                validateCellRangeList(
                            ::std::vector< ::com::sun::star::table::CellRangeAddress >& orRanges,
                            bool bTrackOverflow );

    /** Tries to convert the passed string to a cell range list.

        @param orRanges  (out-parameter) Returns the converted cell range
            addresses. If a range in the passed string contains cells outside
            the currently supported spreadsheet limits, it will be cropped to
            these limits. Example: the range string "A1:ZZ100000" may be
            converted to the range A1:IV65536. If a range is completely outside
            the limits, it will be omitted.
        @param rString  Cell range list string in A1 notation, space separated.
        @param nSheet  Sheet index to be inserted into orRanges (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original ranges contain cells outside of the supported sheet
            limits.
     */
    void                convertToCellRangeList(
                            ::std::vector< ::com::sun::star::table::CellRangeAddress >& orRanges,
                            const ::rtl::OUString& rString,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    /** Tries to convert the passed BIFF range list to a cell range list.

        @param orRanges  (out-parameter) Returns the converted cell range
            addresses. If a range in the passed string contains cells outside
            the currently supported spreadsheet limits, it will be cropped to
            these limits. Example: the range string "A1:ZZ100000" may be
            converted to the range A1:IV65536. If a range is completely outside
            the limits, it will be omitted.
        @param rBiffRanges  List of BIFF cell range objects.
        @param nSheet  Sheet index to be inserted into orRanges (will be checked).
        @param bTrackOverflow  true = Update the internal overflow flags, if
            the original ranges contain cells outside of the supported sheet
            limits.
     */
    void                convertToCellRangeList(
                            ::std::vector< ::com::sun::star::table::CellRangeAddress >& orRanges,
                            const BiffRangeList& rBiffRanges,
                            sal_Int16 nSheet,
                            bool bTrackOverflow );

    // ------------------------------------------------------------------------

    /** Creates a unique key from the passed address. Used for shared formulas
        and table operations in BIFF import, uses column/row only. */
    static sal_Int32    convertToId( const ::com::sun::star::table::CellAddress& rAddress );

    // ------------------------------------------------------------------------
private:
    void                initializeMaxPos(
                            sal_Int16 nMaxXlsTab, sal_Int32 nMaxXlsCol, sal_Int32 nMaxXlsRow );

    void                initializeEncodedUrl(
                            sal_Unicode cUrlThisWorkbook, sal_Unicode cUrlExternal,
                            sal_Unicode cUrlThisSheet, sal_Unicode cUrlInternal );

private:
    ::com::sun::star::table::CellAddress maMaxApiPos;   /// Maximum valid cell address in Calc.
    ::com::sun::star::table::CellAddress maMaxXlsPos;   /// Maximum valid cell address in Excel.
    ::com::sun::star::table::CellAddress maMaxPos;      /// Maximum valid cell address in Calc/Excel.
    sal_Unicode         mcUrlThisWorkbook;              /// Control character: Link to current workbook.
    sal_Unicode         mcUrlExternal;                  /// Control character: Link to external workbook/sheet.
    sal_Unicode         mcUrlThisSheet;                 /// Control character: Link to current sheet.
    sal_Unicode         mcUrlInternal;                  /// Control character: Link to internal sheet.
    bool                mbColOverflow;                  /// Flag for "columns overflow".
    bool                mbRowOverflow;                  /// Flag for "rows overflow".
    bool                mbTabOverflow;                  /// Flag for "tables overflow".
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

