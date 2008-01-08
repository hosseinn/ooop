/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffhelper.hxx,v $
 *
 *  $Revision: 1.1.2.19 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/23 14:15:58 $
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

#ifndef OOX_XLS_BIFFHELPER_HXX
#define OOX_XLS_BIFFHELPER_HXX

#include "oox/core/helper.hxx"

namespace oox {
namespace xls {

class BiffInputStream;
class BiffOutputStream;

// ============================================================================

/** An enumeration for all binary Excel file format types (BIFF types). */
enum BiffType
{
    BIFF2 = 0,                  /// MS Excel 2.1.
    BIFF3,                      /// MS Excel 3.0.
    BIFF4,                      /// MS Excel 4.0.
    BIFF5,                      /// MS Excel 5.0, MS Excel 7.0 (95).
    BIFF8,                      /// MS Excel 8.0 (97), 9.0 (2000), 10.0 (XP), 11.0 (2003).
    BIFF_UNKNOWN                /// Unknown BIFF version.
};

/** An enumeration for all types of fragments in a BIFF workbook stream. */
enum BiffFragmentType
{
    BIFF_FRAGMENT_GLOBALS,      /// Workbook globals fragment.
    BIFF_FRAGMENT_WORKSHEET,    /// Worksheet fragment.
    BIFF_FRAGMENT_CHART,        /// Chart fragment.
    BIFF_FRAGMENT_MACRO,        /// Macro sheet fragment.
    BIFF_FRAGMENT_EMPTYSHEET,   /// Sheet fragment of unsupported type.
    BIFF_FRAGMENT_WORKSPACE,    /// BIFF4 workspace/workbook globals.
    BIFF_FRAGMENT_UNKNOWN       /// Unknown fragment/error.
};

// record identifiers ---------------------------------------------------------

const sal_uInt16 BIFF2_ID_BOF               = 0x0009;
const sal_uInt16 BIFF3_ID_BOF               = 0x0209;
const sal_uInt16 BIFF4_ID_BOF               = 0x0409;
const sal_uInt16 BIFF5_ID_BOF               = 0x0809;
const sal_uInt16 BIFF_ID_CONT               = 0x003C;
const sal_uInt16 BIFF_ID_EOF                = 0x000A;
const sal_uInt16 BIFF_ID_UNKNOWN            = SAL_MAX_UINT16;

// record constants -----------------------------------------------------------

const sal_uInt16 BIFF_BOF_BIFF2             = 0x0200;
const sal_uInt16 BIFF_BOF_BIFF3             = 0x0300;
const sal_uInt16 BIFF_BOF_BIFF4             = 0x0400;
const sal_uInt16 BIFF_BOF_BIFF5             = 0x0500;
const sal_uInt16 BIFF_BOF_BIFF8             = 0x0600;

const sal_uInt8 BIFF_ERR_NULL               = 0x00;
const sal_uInt8 BIFF_ERR_DIV0               = 0x07;
const sal_uInt8 BIFF_ERR_VALUE              = 0x0F;
const sal_uInt8 BIFF_ERR_REF                = 0x17;
const sal_uInt8 BIFF_ERR_NAME               = 0x1D;
const sal_uInt8 BIFF_ERR_NUM                = 0x24;
const sal_uInt8 BIFF_ERR_NA                 = 0x2A;

const sal_uInt8 BIFF_DATATYPE_EMPTY         = 0x00;
const sal_uInt8 BIFF_DATATYPE_DOUBLE        = 0x01;
const sal_uInt8 BIFF_DATATYPE_STRING        = 0x02;
const sal_uInt8 BIFF_DATATYPE_BOOL          = 0x04;
const sal_uInt8 BIFF_DATATYPE_ERROR         = 0x10;

// unicode strings ------------------------------------------------------------

const sal_uInt8 BIFF_STRF_16BIT             = 0x01;
const sal_uInt8 BIFF_STRF_PHONETIC          = 0x04;
const sal_uInt8 BIFF_STRF_RICH              = 0x08;
const sal_uInt8 BIFF_STRF_UNKNOWN           = 0xF2;

/** Flags used to specify import/export mode of strings. */
typedef sal_Int32 BiffStringFlags;

const BiffStringFlags BIFF_STR_DEFAULT      = 0x0000;   /// Default string settings.
const BiffStringFlags BIFF_STR_FORCEUNICODE = 0x0001;   /// Always use UCS-2 characters (default: try to compress). BIFF8 export only.
const BiffStringFlags BIFF_STR_8BITLENGTH   = 0x0002;   /// 8-bit string length field (default: 16-bit).
const BiffStringFlags BIFF_STR_SMARTFLAGS   = 0x0004;   /// Omit flags on empty string (default: read/write always). BIFF8 only.
const BiffStringFlags BIFF_STR_KEEPFONTS    = 0x0008;   /// Keep old fonts when reading unformatted string (default: clear fonts). Import only.
const BiffStringFlags BIFF_STR_EXTRAFONTS   = 0x0010;   /// Read trailing rich-string font array (default: nothing). BIFF2-BIFF5 import only.

// GUID =======================================================================

/** This struct stores a GUID (class ID) and supports reading, writing and comparison.
 */
struct BiffGuid
{
    sal_uInt8           mpnData[ 16 ];      /// Stores the GUID, always in little-endian.

    explicit            BiffGuid();
    explicit            BiffGuid(
                            sal_uInt32 nData1,
                            sal_uInt16 nData2, sal_uInt16 nData3,
                            sal_uInt8 nData41, sal_uInt8 nData42,
                            sal_uInt8 nData43, sal_uInt8 nData44,
                            sal_uInt8 nData45, sal_uInt8 nData46,
                            sal_uInt8 nData47, sal_uInt8 nData48 );
};

bool operator==( const BiffGuid& rGuid1, const BiffGuid& rGuid2 );
bool operator<( const BiffGuid& rGuid1, const BiffGuid& rGuid2 );

BiffInputStream& operator>>( BiffInputStream& rStrm, BiffGuid& rGuid );
BiffOutputStream& operator<<( BiffOutputStream& rStrm, const BiffGuid& rGuid );

// ============================================================================

/** Static helper functions for BIFF filters. */
class BiffHelper
{
public:
    static const BiffGuid maGuidStdHlink;       /// GUID of StdHlink (HLINK record).
    static const BiffGuid maGuidUrlMoniker;     /// GUID of URL moniker (HLINK record).
    static const BiffGuid maGuidFileMoniker;    /// GUID of file moniker (HLINK record).

    // stream -----------------------------------------------------------------

    /** @return  True = passed record identifier is a BOF record. */
    static bool         isBofRecord( sal_uInt16 nRecId );

    /** Skips the current fragment (processes embedded BOF/EOF blocks correctly).

        Skips all records until next EOF record. When this function returns,
        stream points to the EOF record, and the next call of startNextRecord()
        at the stream will start the record following the EOF record.

        @return  True = stream points to the EOF record of the current fragment.
     */
    static bool         skipFragment( BiffInputStream& rStrm );

    /** Starts a new fragment in a workbbok stream and returns the fragment type.

        The passed stream must point before a BOF record. The function will
        try to start the next record and read the contents of the BOF record,
        if extant.

        @return  Fragment type according to the imported BOF record.
     */
    static BiffFragmentType startFragment( BiffInputStream& rStrm, BiffType eBiff );

    // conversion -------------------------------------------------------------

    /** Converts the passed packed number to a double. */
    static double       calcDoubleFromRk( sal_Int32 nRkValue );
    /** Converts the passed double to a packed number, returns true on success. */
    static bool         calcRkFromDouble( sal_Int32& ornRkValue, double fValue );

    /** Converts the passed error to a double containing the respective Calc error code. */
    static double       calcDoubleFromError( sal_uInt8 nErrorCode );

    /** Returns a text encoding from an Windows code page.
        @return  The corresponding text encoding or RTL_TEXTENCODING_DONTKNOW. */
    static rtl_TextEncoding calcTextEncodingFromCodePage( sal_uInt16 nCodePage );
    /** Returns a Windows code page from a text encoding. */
    static sal_uInt16   calcCodePageFromTextEncoding( rtl_TextEncoding eTextEnc );
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

