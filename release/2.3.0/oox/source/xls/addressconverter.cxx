/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: addressconverter.cxx,v $
 *
 *  $Revision: 1.1.2.19 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/03 13:49:35 $
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

#include "oox/xls/addressconverter.hxx"
#include <osl/diagnose.h>
#include <rtl/ustrbuf.hxx>
#include <rtl/strbuf.hxx>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/sheet/XCellRangeAddressable.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/biffoutputstream.hxx"

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::rtl::OStringBuffer;
using ::rtl::OUStringToOString;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::container::XIndexAccess;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::sheet::XCellRangeAddressable;

namespace oox {
namespace xls {

// ============================================================================

//! TODO: this limit may be changed
const sal_Int16 API_MAXTAB          = 255;

const sal_Int32 OOX_MAXCOL          = static_cast< sal_Int32 >( (1 << 14) - 1 );
const sal_Int32 OOX_MAXROW          = static_cast< sal_Int32 >( (1 << 20) - 1 );
const sal_Int16 OOX_MAXTAB          = static_cast< sal_Int16 >( (1 << 15) - 1 );

const sal_Int32 BIFF2_MAXCOL        = 255;
const sal_Int32 BIFF2_MAXROW        = 16383;
const sal_Int16 BIFF2_MAXTAB        = 0;

const sal_Int32 BIFF3_MAXCOL        = BIFF2_MAXCOL;
const sal_Int32 BIFF3_MAXROW        = BIFF2_MAXROW;
const sal_Int16 BIFF3_MAXTAB        = BIFF2_MAXTAB;

const sal_Int32 BIFF4_MAXCOL        = BIFF3_MAXCOL;
const sal_Int32 BIFF4_MAXROW        = BIFF3_MAXROW;
const sal_Int16 BIFF4_MAXTAB        = 32767;

const sal_Int32 BIFF5_MAXCOL        = BIFF4_MAXCOL;
const sal_Int32 BIFF5_MAXROW        = BIFF4_MAXROW;
const sal_Int16 BIFF5_MAXTAB        = BIFF4_MAXTAB;

const sal_Int32 BIFF8_MAXCOL        = BIFF5_MAXCOL;
const sal_Int32 BIFF8_MAXROW        = 65535;
const sal_Int16 BIFF8_MAXTAB        = BIFF5_MAXTAB;

const sal_Unicode BIFF_URL_DRIVE    = '\x01';       /// DOS drive letter or UNC path.
const sal_Unicode BIFF_URL_ROOT     = '\x02';       /// Root directory of current drive.
const sal_Unicode BIFF_URL_SUBDIR   = '\x03';       /// Subdirectory delimiter.
const sal_Unicode BIFF_URL_PARENT   = '\x04';       /// Parent directory.
const sal_Unicode BIFF_URL_RAW      = '\x05';       /// Unencoded URL.
const sal_Unicode BIFF_URL_INSTALL  = '\x06';       /// Application installation directory.
const sal_Unicode BIFF_URL_INSTALL2 = '\x07';       /// Alternative application installation directory.
const sal_Unicode BIFF_URL_ADDIN    = '\x08';       /// Add-in installation directory.
const sal_Unicode BIFF4_URL_SHEET   = '\x09';       /// BIFF4 internal sheet.
const sal_Unicode BIFF_URL_UNC      = '@';          /// UNC path root.

// ============================================================================
// ============================================================================

void BiffAddress::read( BiffInputStream& rStrm, bool bCol16Bit )
{
    rStrm >> mnRow;
    if( bCol16Bit )
        rStrm >> mnCol;
    else
        mnCol = rStrm.readuInt8();
}

void BiffAddress::write( BiffOutputStream& rStrm, bool bCol16Bit ) const
{
    rStrm << mnRow;
    if( bCol16Bit )
        rStrm << mnCol;
    else
        rStrm << static_cast< sal_uInt8 >( mnCol );
}

bool operator==( const BiffAddress& rL, const BiffAddress& rR )
{
    return (rL.mnCol == rR.mnCol) && (rL.mnRow == rR.mnRow);
}

bool operator<( const BiffAddress& rL, const BiffAddress& rR )
{
    return (rL.mnCol < rR.mnCol) || ((rL.mnCol == rR.mnCol) && (rL.mnRow < rR.mnRow));
}

BiffInputStream& operator>>( BiffInputStream& rStrm, BiffAddress& rPos )
{
    rPos.read( rStrm );
    return rStrm;
}

BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffAddress& rPos )
{
    rPos.write( rStrm );
    return rStrm;
}

// ----------------------------------------------------------------------------

bool BiffRange::contains( const BiffAddress& rPos ) const
{
    return  (maFirst.mnCol <= rPos.mnCol) && (rPos.mnCol <= maLast.mnCol) &&
            (maFirst.mnRow <= rPos.mnRow) && (rPos.mnRow <= maLast.mnRow);
}

void BiffRange::read( BiffInputStream& rStrm, bool bCol16Bit )
{
    rStrm >> maFirst.mnRow >> maLast.mnRow;
    if( bCol16Bit )
        rStrm >> maFirst.mnCol >> maLast.mnCol;
    else
    {
        maFirst.mnCol = rStrm.readuInt8();
        maLast.mnCol = rStrm.readuInt8();
    }
}

void BiffRange::write( BiffOutputStream& rStrm, bool bCol16Bit ) const
{
    rStrm << maFirst.mnRow << maLast.mnRow;
    if( bCol16Bit )
        rStrm << maFirst.mnCol << maLast.mnCol;
    else
        rStrm << static_cast< sal_uInt8 >( maFirst.mnCol ) << static_cast< sal_uInt8 >( maLast.mnCol );
}

bool operator==( const BiffRange& rL, const BiffRange& rR )
{
    return (rL.maFirst == rR.maFirst) && (rL.maLast == rR.maLast);
}

bool operator<( const BiffRange& rL, const BiffRange& rR )
{
    return (rL.maFirst < rR.maFirst) || ((rL.maFirst == rR.maFirst) && (rL.maLast < rR.maLast));
}

BiffInputStream& operator>>( BiffInputStream& rStrm, BiffRange& rRange )
{
    rRange.read( rStrm );
    return rStrm;
}

BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffRange& rRange )
{
    rRange.write( rStrm );
    return rStrm;
}

// ----------------------------------------------------------------------------

BiffRange BiffRangeList::getEnclosingRange() const
{
    BiffRange aRange;
    if( !empty() )
    {
        const_iterator aIt = begin(), aEnd = end();
        aRange = *aIt;
        for( ++aIt; aIt != aEnd; ++aIt )
        {
            aRange.maFirst.mnCol = ::std::min( aRange.maFirst.mnCol, aIt->maFirst.mnCol );
            aRange.maFirst.mnRow = ::std::min( aRange.maFirst.mnRow, aIt->maFirst.mnRow );
            aRange.maLast.mnCol  = ::std::max( aRange.maLast.mnCol, aIt->maLast.mnCol );
            aRange.maLast.mnRow  = ::std::max( aRange.maLast.mnRow, aIt->maLast.mnRow );
        }
    }
    return aRange;
}

void BiffRangeList::read( BiffInputStream& rStrm, bool bCol16Bit )
{
    sal_uInt16 nCount;
    rStrm >> nCount;
    reserve( size() + nCount );
    for( sal_uInt16 nIndex = 0; rStrm.isValid() && (nIndex < nCount); ++nIndex )
    {
        resize( size() + 1 );
        back().read( rStrm, bCol16Bit );
    }
}

void BiffRangeList::write( BiffOutputStream& rStrm, bool bCol16Bit ) const
{
    writeSubList( rStrm, 0, size(), bCol16Bit );
}

void BiffRangeList::writeSubList( BiffOutputStream& rStrm, size_t nBegin, size_t nCount, bool bCol16Bit ) const
{
    OSL_ENSURE( nBegin <= size(), "BiffRangeList::writeSubList - invalid start position" );
    size_t nEnd = ::std::min< size_t >( nBegin + nCount, size() );
    sal_uInt16 nBiffCount = getLimitedValue< sal_uInt16, size_t >( nEnd - nBegin, 0, SAL_MAX_UINT16 );
    rStrm << nBiffCount;
    rStrm.setSliceSize( bCol16Bit ? 8 : 6 );
    for( const_iterator aIt = begin() + nBegin, aEnd = begin() + nEnd; aIt != aEnd; ++aIt )
        aIt->write( rStrm, bCol16Bit );
}

BiffInputStream& operator>>( BiffInputStream& rStrm, BiffRangeList& rRanges )
{
    rRanges.read( rStrm );
    return rStrm;
}

BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffRangeList& rRanges )
{
    rRanges.write( rStrm );
    return rStrm;
}

// ============================================================================

AddressConverter::AddressConverter( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mcUrlThisWorkbook( 0 ),
    mcUrlExternal( 0 ),
    mcUrlThisSheet( 0 ),
    mcUrlInternal( 0 ),
    mbColOverflow( false ),
    mbRowOverflow( false ),
    mbTabOverflow( false )
{
    switch( getFilterType() )
    {
        case FILTER_OOX:
            initializeMaxPos( OOX_MAXTAB, OOX_MAXCOL, OOX_MAXROW );
        break;
        case FILTER_BIFF: switch( getBiff() )
        {
            case BIFF2:
                initializeMaxPos( BIFF2_MAXTAB, BIFF2_MAXCOL, BIFF2_MAXROW );
                initializeEncodedUrl( '\x00', '\x01', '\x02', '\x00' );
            break;
            case BIFF3:
                initializeMaxPos( BIFF3_MAXTAB, BIFF3_MAXCOL, BIFF3_MAXROW );
                initializeEncodedUrl( '\x00', '\x01', '\x02', '\x00' );
            break;
            case BIFF4:
                initializeMaxPos( BIFF4_MAXTAB, BIFF4_MAXCOL, BIFF4_MAXROW );
                initializeEncodedUrl( '\x00', '\x01', '\x02', '\x00' );
            break;
            case BIFF5:
                initializeMaxPos( BIFF5_MAXTAB, BIFF5_MAXCOL, BIFF5_MAXROW );
                initializeEncodedUrl( '\x04', '\x01', '\x02', '\x03' );
            break;
            case BIFF8:
                initializeMaxPos( BIFF8_MAXTAB, BIFF8_MAXCOL, BIFF8_MAXROW );
                initializeEncodedUrl( '\x04', '\x01', '\x00', '\x02' );
            break;
            case BIFF_UNKNOWN: break;
        }
        break;
        case FILTER_UNKNOWN: break;
    }
}

// ----------------------------------------------------------------------------

bool AddressConverter::parseAddress2d(
        sal_Int32& ornColumn, sal_Int32& ornRow,
        const OUString& rString, sal_Int32 nStart, sal_Int32 nLength )
{
    ornColumn = ornRow = 0;
    if( (nStart < 0) || (nStart >= rString.getLength()) || (nLength < 2) )
        return false;

    const sal_Unicode* pcChar = rString.getStr() + nStart;
    const sal_Unicode* pcEndChar = pcChar + ::std::min( nLength, rString.getLength() - nStart );

    enum { STATE_COL, STATE_ROW } eState = STATE_COL;
    while( pcChar < pcEndChar )
    {
        sal_Unicode cChar = *pcChar;
        switch( eState )
        {
            case STATE_COL:
            {
                if( ('a' <= cChar) && (cChar <= 'z') )
                    (cChar -= 'a') += 'A';
                if( ('A' <= cChar) && (cChar <= 'Z') )
                {
                    /*  Return, if 1-based column index is already 6 characters
                        long (12356631 is column index for column AAAAAA). */
                    if( ornColumn >= 12356631 )
                        return false;
                    (ornColumn *= 26) += (cChar - 'A' + 1);
                }
                else if( ornColumn > 0 )
                {
                    --pcChar;
                    eState = STATE_ROW;
                }
                else
                    return false;
            }
            break;

            case STATE_ROW:
            {
                if( ('0' <= cChar) && (cChar <= '9') )
                {
                    // return, if 1-based row is already 9 digits long
                    if( ornRow >= 100000000 )
                        return false;
                    (ornRow *= 10) += (cChar - '0');
                }
                else
                    return false;
            }
            break;
        }
        ++pcChar;
    }

    --ornColumn;
    --ornRow;
    return (ornColumn >= 0) && (ornRow >= 0);
}

bool AddressConverter::parseRange2d(
        sal_Int32& ornStartColumn, sal_Int32& ornStartRow,
        sal_Int32& ornEndColumn, sal_Int32& ornEndRow,
        const OUString& rString, sal_Int32 nStart, sal_Int32 nLength )
{
    ornStartColumn = ornStartRow = ornEndColumn = ornEndRow = 0;
    if( (nStart < 0) || (nStart >= rString.getLength()) || (nLength < 2) )
        return false;

    sal_Int32 nEnd = nStart + ::std::min( nLength, rString.getLength() - nStart );
    sal_Int32 nColonPos = rString.indexOf( ':', nStart );
    if( (nStart < nColonPos) && (nColonPos + 1 < nEnd) )
    {
        return
            parseAddress2d( ornStartColumn, ornStartRow, rString, nStart, nColonPos - nStart ) &&
            parseAddress2d( ornEndColumn, ornEndRow, rString, nColonPos + 1, nLength - nColonPos - 1 );
    }

    if( parseAddress2d( ornStartColumn, ornStartRow, rString, nStart, nLength ) )
    {
        ornEndColumn = ornStartColumn;
        ornEndRow = ornStartRow;
        return true;
    }

    return false;
}

namespace {

bool lclAppendUrlChar( OUStringBuffer& orUrl, sal_Unicode cChar, bool bEncodeSpecial )
{
    // #126855# encode special characters
    if( bEncodeSpecial ) switch( cChar )
    {
        case '#':   orUrl.appendAscii( "%23" );  return true;
        case '%':   orUrl.appendAscii( "%25" );  return true;
    }
    orUrl.append( cChar );
    return cChar >= ' ';
}

} // namespace

bool AddressConverter::parseEncodedTarget(
        OUString& orClassName, OUString& orTargetUrl, OUString& orSheetName,
        const OUString& rBiffEncoded )
{
    OUStringBuffer aTargetUrl;
    OUStringBuffer aSheetName;

    enum
    {
        STATE_START,
        STATE_ENCODED_PATH_START,       /// Start of encoded file path.
        STATE_ENCODED_PATH,             /// Inside encoded file path.
        STATE_ENCODED_DRIVE,            /// DOS drive letter or start of UNC path.
        STATE_ENCODED_URL,              /// Encoded URL, e.g. http links.
        STATE_UNENCODED,                /// Unencoded URL, could be DDE or OLE.
        STATE_DDE_OLE,                  /// Second part of DDE or OLE link.
        STATE_FILENAME,                 /// File name enclosed in brackets.
        STATE_SHEETNAME,                /// Sheet name following enclosed file name.
        STATE_UNSUPPORTED,              /// Unsupported special paths.
        STATE_ERROR
    }
    eState = STATE_START;

    const sal_Unicode* pcChar = rBiffEncoded.getStr();
    const sal_Unicode* pcEnd = pcChar + rBiffEncoded.getLength();
    for( ; (eState != STATE_ERROR) && (pcChar < pcEnd) && (*pcChar != 0); ++pcChar )
    {
        sal_Unicode cChar = *pcChar;
        switch( eState )
        {
            case STATE_START:
                if( (cChar == mcUrlThisWorkbook) || (cChar == mcUrlThisSheet) )
                {
                    if( pcChar + 1 < pcEnd ) eState = STATE_ERROR;
                }
                else if( cChar == mcUrlExternal )
                    eState = (pcChar + 1 < pcEnd) ? STATE_ENCODED_PATH_START : STATE_ERROR;
                else if( cChar == mcUrlInternal )
                    eState = (pcChar + 1 < pcEnd) ? STATE_SHEETNAME : STATE_ERROR;
                else
                    eState = lclAppendUrlChar( aTargetUrl, cChar, true ) ? STATE_UNENCODED : STATE_ERROR;
            break;

            case STATE_ENCODED_PATH_START:
                if( cChar == BIFF_URL_DRIVE )
                    eState = STATE_ENCODED_DRIVE;
                else if( cChar == BIFF_URL_ROOT )
                {
                    aTargetUrl.append( sal_Unicode( '/' ) );
                    eState = STATE_ENCODED_PATH;
                }
                else if( cChar == BIFF_URL_PARENT )
                    aTargetUrl.appendAscii( "../" );
                else if( cChar == BIFF_URL_RAW )
                    eState = STATE_ENCODED_URL;
                else if( cChar == BIFF_URL_INSTALL )
                    eState = STATE_UNSUPPORTED;
                else if( cChar == BIFF_URL_INSTALL2 )
                    eState = STATE_UNSUPPORTED;
                else if( cChar == BIFF_URL_ADDIN )
                    eState = STATE_UNSUPPORTED;
                else if( (getBiff() == BIFF4) && (cChar == BIFF4_URL_SHEET) )
                    eState = STATE_SHEETNAME;
                else if( lclAppendUrlChar( aTargetUrl, cChar, true ) )
                    eState = STATE_ENCODED_PATH;
                else
                    eState = STATE_ERROR;
            break;

            case STATE_ENCODED_PATH:
                if( cChar == BIFF_URL_SUBDIR )
                    aTargetUrl.append( sal_Unicode( '/' ) );
                else if( cChar == '[' )
                    eState = STATE_FILENAME;
                else if( !lclAppendUrlChar( aTargetUrl, cChar, true ) )
                    eState = STATE_ERROR;
            break;

            case STATE_ENCODED_DRIVE:
                if( cChar == BIFF_URL_UNC )
                {
                    aTargetUrl.appendAscii( "file://" );
                    eState = STATE_ENCODED_PATH;
                }
                else
                {
                    aTargetUrl.appendAscii( "file:///" );
                    eState = lclAppendUrlChar( aTargetUrl, cChar, false ) ? STATE_ENCODED_PATH : STATE_ERROR;
                    aTargetUrl.appendAscii( ":/" );
                }
            break;

            case STATE_ENCODED_URL:
            {
                sal_Int32 nLength = cChar;
                if( nLength + 1 == pcEnd - pcChar )
                    aTargetUrl.append( pcChar + 1, nLength );
                else
                    eState = STATE_ERROR;
            }
            break;

            case STATE_UNENCODED:
                if( cChar == BIFF_URL_SUBDIR )
                {
                    orClassName = aTargetUrl.makeStringAndClear();
                    eState = STATE_DDE_OLE;
                }
                else if( cChar == '[' )
                    eState = STATE_FILENAME;
                else if( !lclAppendUrlChar( aTargetUrl, cChar, true ) )
                    eState = STATE_ERROR;
            break;

            case STATE_DDE_OLE:
                if( !lclAppendUrlChar( aTargetUrl, cChar, true ) )
                    eState = STATE_ERROR;
            break;

            case STATE_FILENAME:
                if( cChar == ']' )
                    eState = STATE_SHEETNAME;
                else if( !lclAppendUrlChar( aTargetUrl, cChar, true ) )
                    eState = STATE_ERROR;
            break;

            case STATE_SHEETNAME:
                if( !lclAppendUrlChar( aSheetName, cChar, false ) )
                    eState = STATE_ERROR;
            break;

            case STATE_UNSUPPORTED:
                pcChar = pcEnd - 1;
            break;

            case STATE_ERROR:
            break;
        }
    }

    orTargetUrl = aTargetUrl.makeStringAndClear();
    orSheetName = aSheetName.makeStringAndClear();

    OSL_ENSURE( (eState != STATE_ERROR) && (pcChar == pcEnd),
        OStringBuffer( "AddressConverter::parseEncodedTarget - parser error in target \"" ).
        append( OUStringToOString( rBiffEncoded, RTL_TEXTENCODING_UTF8 ) ).append( '"' ).getStr() );
    return (eState != STATE_ERROR) && (eState != STATE_UNSUPPORTED) && (pcChar == pcEnd);
}

// ----------------------------------------------------------------------------

bool AddressConverter::checkCol( sal_Int32 nCol, bool bTrackOverflow )
{
    bool bValid = (0 <= nCol) && (nCol <= maMaxPos.Column);
    if( !bValid && bTrackOverflow )
        mbColOverflow = true;
    return bValid;
}

bool AddressConverter::checkRow( sal_Int32 nRow, bool bTrackOverflow )
{
    bool bValid = (0 <= nRow) && (nRow <= maMaxPos.Row);
    if( !bValid && bTrackOverflow )
        mbRowOverflow = true;
    return bValid;
}

bool AddressConverter::checkTab( sal_Int16 nSheet, bool bTrackOverflow )
{
    bool bValid = (0 <= nSheet) && (nSheet <= maMaxPos.Sheet);
    if( !bValid && bTrackOverflow )
        mbTabOverflow |= (nSheet > maMaxPos.Sheet);  // do not warn for deleted refs (-1)
    return bValid;
}

// ----------------------------------------------------------------------------

bool AddressConverter::checkCellAddress( const CellAddress& rAddress, bool bTrackOverflow )
{
    return
        checkTab( rAddress.Sheet, bTrackOverflow ) &&
        checkCol( rAddress.Column, bTrackOverflow ) &&
        checkRow( rAddress.Row, bTrackOverflow );
}

bool AddressConverter::convertToCellAddress( CellAddress& orAddress,
        const OUString& rString, sal_Int16 nSheet, bool bTrackOverflow )
{
    orAddress.Sheet = nSheet;
    return
        parseAddress2d( orAddress.Column, orAddress.Row, rString ) &&
        checkCellAddress( orAddress, bTrackOverflow );
}

bool AddressConverter::convertToCellAddress( CellAddress& orAddress,
        const BiffAddress& rBiffAddress, sal_Int16 nSheet, bool bTrackOverflow )
{
    orAddress.Sheet = nSheet;
    orAddress.Column = static_cast< sal_Int32 >( rBiffAddress.mnCol );
    orAddress.Row = static_cast< sal_Int32 >( rBiffAddress.mnRow );
    return checkCellAddress( orAddress, bTrackOverflow );
}

CellAddress AddressConverter::createValidCellAddress(
        const OUString& rString, sal_Int16 nSheet, bool bTrackOverflow )
{
    CellAddress aAddress;
    if( !convertToCellAddress( aAddress, rString, nSheet, bTrackOverflow ) )
    {
        aAddress.Sheet  = getLimitedValue< sal_Int16, sal_Int16 >( nSheet, 0, maMaxPos.Sheet );
        aAddress.Column = ::std::min( aAddress.Column, maMaxPos.Column );
        aAddress.Row    = ::std::min( aAddress.Row, maMaxPos.Row );
    }
    return aAddress;
}

CellAddress AddressConverter::createValidCellAddress(
        const BiffAddress& rBiffAddress, sal_Int16 nSheet, bool bTrackOverflow )
{
    CellAddress aAddress;
    if( !convertToCellAddress( aAddress, rBiffAddress, nSheet, bTrackOverflow ) )
    {
        aAddress.Sheet  = getLimitedValue< sal_Int16, sal_Int16 >( nSheet, 0, maMaxPos.Sheet );
        aAddress.Column = ::std::min( static_cast< sal_Int32 >( rBiffAddress.mnCol ), maMaxPos.Column );
        aAddress.Row    = ::std::min( static_cast< sal_Int32 >( rBiffAddress.mnRow ), maMaxPos.Row );
    }
    return aAddress;
}

// ----------------------------------------------------------------------------

bool AddressConverter::checkCellRange( const CellRangeAddress& rRange, bool bTrackOverflow )
{
    checkCol( rRange.EndColumn, bTrackOverflow );
    checkRow( rRange.EndRow, bTrackOverflow );
    return
        checkTab( rRange.Sheet, bTrackOverflow ) &&
        checkCol( rRange.StartColumn, bTrackOverflow ) &&
        checkRow( rRange.StartRow, bTrackOverflow );
}

bool AddressConverter::validateCellRange( CellRangeAddress& orRange, bool bTrackOverflow )
{
    if( orRange.StartColumn > orRange.EndColumn )
        ::std::swap( orRange.StartColumn, orRange.EndColumn );
    if( orRange.StartRow > orRange.EndRow )
        ::std::swap( orRange.StartRow, orRange.EndRow );
    if( !checkCellRange( orRange, bTrackOverflow ) )
        return false;
    if( orRange.EndColumn > maMaxPos.Column )
        orRange.EndColumn = maMaxPos.Column;
    if( orRange.EndRow > maMaxPos.Row )
        orRange.EndRow = maMaxPos.Row;
    return true;
}

bool AddressConverter::convertToCellRange( CellRangeAddress& orRange,
        const OUString& rString, sal_Int16 nSheet, bool bTrackOverflow )
{
    orRange.Sheet = nSheet;
    return
        parseRange2d( orRange.StartColumn, orRange.StartRow, orRange.EndColumn, orRange.EndRow, rString ) &&
        validateCellRange( orRange, bTrackOverflow );
}

bool AddressConverter::convertToCellRange( CellRangeAddress& orRange,
        const BiffRange& rBiffRange, sal_Int16 nSheet, bool bTrackOverflow )
{
    orRange.Sheet = nSheet;
    orRange.StartColumn = static_cast< sal_Int32 >( rBiffRange.maFirst.mnCol );
    orRange.StartRow = static_cast< sal_Int32 >( rBiffRange.maFirst.mnRow );
    orRange.EndColumn = static_cast< sal_Int32 >( rBiffRange.maLast.mnCol );
    orRange.EndRow = static_cast< sal_Int32 >( rBiffRange.maLast.mnRow );
    return validateCellRange( orRange, bTrackOverflow );
}

// ----------------------------------------------------------------------------

bool AddressConverter::checkCellRangeList( const ::std::vector< CellRangeAddress >& rRanges, bool bTrackOverflow )
{
    for( ::std::vector< CellRangeAddress >::const_iterator aIt = rRanges.begin(), aEnd = rRanges.end(); aIt != aEnd; ++aIt )
        if( !checkCellRange( *aIt, bTrackOverflow ) )
            return false;
    return true;
}

void AddressConverter::validateCellRangeList( ::std::vector< CellRangeAddress >& orRanges, bool bTrackOverflow )
{
    for( size_t nIndex = orRanges.size(); nIndex > 0; --nIndex )
        if( !validateCellRange( orRanges[ nIndex - 1 ], bTrackOverflow ) )
            orRanges.erase( orRanges.begin() + nIndex - 1 );
}

void AddressConverter::convertToCellRangeList( ::std::vector< CellRangeAddress >& orRanges,
        const OUString& rString, sal_Int16 nSheet, bool bTrackOverflow )
{
    sal_Int32 nPos = 0;
    sal_Int32 nLen = rString.getLength();
    CellRangeAddress aRange;
    while( (0 <= nPos) && (nPos < nLen) )
    {
        OUString aToken = rString.getToken( 0, ' ', nPos );
        if( (aToken.getLength() > 0) && convertToCellRange( aRange, aToken, nSheet, bTrackOverflow ) )
            orRanges.push_back( aRange );
    }
}

void AddressConverter::convertToCellRangeList( ::std::vector< CellRangeAddress >& orRanges,
        const BiffRangeList& rBiffRanges, sal_Int16 nSheet, bool bTrackOverflow )
{
    CellRangeAddress aRange;
    for( BiffRangeList::const_iterator aIt = rBiffRanges.begin(), aEnd = rBiffRanges.end(); aIt != aEnd; ++aIt )
        if( convertToCellRange( aRange, *aIt, nSheet, bTrackOverflow ) )
            orRanges.push_back( aRange );
}

sal_Int32 AddressConverter::convertToId( const CellAddress& rAddress )
{
    return (rAddress.Column << 16) | rAddress.Row;
}

// private --------------------------------------------------------------------

void AddressConverter::initializeMaxPos(
        sal_Int16 nMaxXlsTab, sal_Int32 nMaxXlsCol, sal_Int32 nMaxXlsRow )
{
    maMaxXlsPos.Sheet  = nMaxXlsTab;
    maMaxXlsPos.Column = nMaxXlsCol;
    maMaxXlsPos.Row    = nMaxXlsRow;

    // maximum cell position in Calc
    try
    {
        Reference< XIndexAccess > xSheetsIA( getDocument()->getSheets(), UNO_QUERY_THROW );
        Reference< XCellRangeAddressable > xAddressable( xSheetsIA->getByIndex( 0 ), UNO_QUERY_THROW );
        CellRangeAddress aRange = xAddressable->getRangeAddress();
        maMaxApiPos = CellAddress( API_MAXTAB, aRange.EndColumn, aRange.EndRow );
        maMaxPos = CellAddress(
            ::std::min( maMaxApiPos.Sheet,  maMaxXlsPos.Sheet ),
            ::std::min( maMaxApiPos.Column, maMaxXlsPos.Column ),
            ::std::min( maMaxApiPos.Row,    maMaxXlsPos.Row ) );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "AddressConverter::AddressConverter - cannot get sheet limits" );
    }
}

void AddressConverter::initializeEncodedUrl(
        sal_Unicode cUrlThisWorkbook, sal_Unicode cUrlExternal,
        sal_Unicode cUrlThisSheet, sal_Unicode cUrlInternal )
{
    mcUrlThisWorkbook   = cUrlThisWorkbook;
    mcUrlExternal       = cUrlExternal;
    mcUrlThisSheet      = cUrlThisSheet;
    mcUrlInternal       = cUrlInternal;
}

// ============================================================================

} // namespace xls
} // namespace oox

