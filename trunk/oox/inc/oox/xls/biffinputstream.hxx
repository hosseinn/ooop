/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffinputstream.hxx,v $
 *
 *  $Revision: 1.1.2.13 $
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

#ifndef OOX_XLS_BIFFINPUTSTREAM_HXX
#define OOX_XLS_BIFFINPUTSTREAM_HXX

#include <memory>
#include "oox/xls/biffhelper.hxx"
#include "oox/xls/biffcodec.hxx"

namespace oox { namespace core {
    class BinaryInputStream;
} }

namespace oox {
namespace xls {

// ============================================================================

const sal_uInt32 BIFF_REC_SEEK_TO_BEGIN     = 0;
const sal_uInt32 BIFF_REC_SEEK_TO_END       = SAL_MAX_UINT32;

const sal_uInt16 BIFF2_MAXRECSIZE           = 2080;
const sal_uInt16 BIFF8_MAXRECSIZE           = 8224;

const sal_Unicode BIFF_DEF_NUL_SUBST_CHAR   = '?';

// ============================================================================

class BiffInputStreamImpl;

/** This class is used to import BIFF streams.

    An instance is constructed with an ::oox::core::BinaryInputStream object.
    The passed stream is reset to its start while constructing this stream.

    To start reading a record call startNextRecord(). Now it is possible to
    read all contents of the record using operator>>() or any of the read***()
    functions. If some data exceeds the record size limit, the stream looks for
    a following CONTINUE record and jumps automatically to it. It is NOT
    allowed that an atomic data type is split into two records (e.g. 4 bytes of
    a double in one record and the other 4 bytes in a following CONTINUE).

    Trying to read over the record limits results in a stream error. The
    isValid() function indicates that by returning false. From now on the data
    returned by the read functions is undefined. The error state will be reset,
    if the record is reset (with the function resetRecord()), or if the next
    record is started.

    To switch off the automatic lookup of CONTINUE records, use resetRecord()
    with false parameter. This is useful e.g. on import of drawing layer data,
    where sometimes solely CONTINUE records will occur. The automatic lookup
    keeps switched off until the method resetRecord() is called with parameter
    true. All other settings done on the stream (e.g. alternative CONTINUE
    record identifier, enabled decryption, NUL substitution character) will be
    reset to default values, if a new record is started.

    The import stream supports decrypting the stream data. The contents of a
    record (not the record header) will be encrypted by Excel if the file has
    been stored with password protection. The functions setDecoder() and
    enableDecoder() control the usage of the decryption algorithms.
    setDecoder() sets�a new decryption algorithm and initially enables it.
    enableDecoder( false ) may be used to stop the usage of the decryption
    temporarily (sometimes record contents are never encrypted, e.g. all BOF
    records or the stream position in BOUNDSHEET). Decryption will be reenabled
    automatically, if a new record is started with the function
    startNextRecord().

    Be careful with used data types:
    sal_uInt16: Record identifiers, raw size of single records.
    sal_uInt32: Record position and size (including CONTINUE records).
    sal_Int64: Core stream position and size.
*/
class BiffInputStream
{
public:
    /** Constructs the Excel record import stream using the passed stream.

        @param rInStream
            The base input stream. Must be seekable. Will be seeked to its
            start position.

        @param bContLookup  Automatic CONTINUE lookup on/off.
     */
    explicit            BiffInputStream(
                            ::oox::core::BinaryInputStream& rInStream,
                            bool bContLookup = true );

                        ~BiffInputStream();

    // record control ---------------------------------------------------------

    /** Sets stream pointer to the start of the next record content.

        Ignores all CONTINUE records of the current record, if automatic
        CONTINUE usage is switched on.

        @return  False = no record found (end of stream).
     */
    bool                startNextRecord();

    /** Sets stream pointer to the start of the content of the specified record.

        The handle of the current record can be received and stored using the
        function getRecHandle() for later usage with this function.

        @return  False = no record found (invalid handle passed).
     */
    bool                startRecordByHandle( sal_Int64 nRecHandle );

    /** Sets stream pointer to begin of record content.

        @param bContLookup
            Automatic CONTINUE lookup on/off. In difference to other stream
            settings, this setting is persistent until next call of this
            function (because it is wanted to receive the next CONTINUE records
            separately).
        @param nAltContId
            Sets an alternative record identifier for content continuation.
            This value is reset automatically when a new record is started with
            startNextRecord().
     */
    void                resetRecord(
                            bool bContLookup,
                            sal_uInt16 nAltContId = BIFF_ID_UNKNOWN );

    /** Sets stream pointer before current record and invalidates stream.

        The next call to startNextRecord() will start again the current record.
        This can be used in situations where a loop or a function leaves on a
        specific record, but the parent context expects to start this record by
        itself. The stream is invalid as long as the first record has not been
        started (it is not allowed to call any other stream operation then).
     */
    void                rewindRecord();

    // decoder ----------------------------------------------------------------

    /** Sets a new decoder object.

        Enables decryption of record contents for the rest of the stream.
     */
    void                setDecoder( BiffDecoderRef xDecoder );

    /** Returns the current decoder object. */
    BiffDecoderRef      getDecoder() const;

    /** Enables/disables usage of current decoder.

        Decryption is reenabled automatically, if a new record is started using
        the function startNextRecord().
     */
    void                enableDecoder( bool bEnable = true );

    // stream/record state and info -------------------------------------------

    /** Returns record reading state: false = record overread. */
    bool                isValid() const;
    /** Returns the current record identifier. */
    sal_uInt16          getRecId() const;
    /** Returns the position inside of the whole record content. */
    sal_uInt32          getRecPos() const;
    /** Returns the data size of the whole record without record headers. */
    sal_uInt32          getRecSize() const;
    /** Returns remaining data size of the whole record without record headers. */
    sal_uInt32          getRecLeft() const;
    /** Returns a unique handle for the current record that can be used with
        the function startRecordByHandle(). */
    sal_Int64           getRecHandle() const;
    /** Returns the record identifier of the following record. */
    sal_uInt16          getNextRecId() const;

    /** Returns the absolute core stream position. */
    sal_Int64           getCoreStreamPos() const;
    /** Returns the stream size. */
    sal_Int64           getCoreStreamSize() const;

    // stream read access -----------------------------------------------------

    /** Reads nBytes bytes to the existing(!) buffer pData.

        @return  Number of bytes really read.
     */
    sal_uInt32          read( void* pData, sal_uInt32 nBytes );

    BiffInputStream&    operator>>( sal_Int8& rnValue );
    BiffInputStream&    operator>>( sal_uInt8& rnValue );
    BiffInputStream&    operator>>( sal_Int16& rnValue );
    BiffInputStream&    operator>>( sal_uInt16& rnValue );
    BiffInputStream&    operator>>( sal_Int32& rnValue );
    BiffInputStream&    operator>>( sal_uInt32& rnValue );
    BiffInputStream&    operator>>( sal_Int64& rnValue );
    BiffInputStream&    operator>>( sal_uInt64& rnValue );
    BiffInputStream&    operator>>( float& rfValue );
    BiffInputStream&    operator>>( double& rfValue );

    sal_Int8            readInt8();
    sal_uInt8           readuInt8();
    sal_Int16           readInt16();
    sal_uInt16          readuInt16();
    sal_Int32           readInt32();
    sal_uInt32          readuInt32();
    sal_Int64           readInt64();
    sal_uInt64          readuInt64();
    float               readFloat();
    double              readDouble();

    // seeking ----------------------------------------------------------------

    /** Seeks absolute in record content to the specified position.

        The value 0 means start of record, independent from physical stream
        position.
     */
    void                seek( sal_uInt32 nRecPos );

    /** Seeks forward inside the current record. */
    void                ignore( sal_uInt32 nBytes );

    // character arrays -------------------------------------------------------

    /** Reads nChar byte characters and returns the string. */
    ::rtl::OString      readCharArray( sal_uInt16 nChars );
    /** Reads nChar byte characters and returns the string. */
    ::rtl::OUString     readCharArray( sal_uInt16 nChars, rtl_TextEncoding eTextEnc );
    /** Reads nChars Unicode characters and returns the string. */
    ::rtl::OUString     readUnicodeArray( sal_uInt16 nChars );

    // byte strings -----------------------------------------------------------

    /** Reads 8/16 bit string length, character array and returns the string. */
    ::rtl::OString      readByteString( bool b16BitLen );
    /** Reads 8/16 bit string length, character array and returns the string. */
    ::rtl::OUString     readByteString( bool b16BitLen, rtl_TextEncoding eTextEnc );

    /** Ignores 8/16 bit string length, character array. */
    void                ignoreByteString( bool b16BitLen );

    // Unicode strings --------------------------------------------------------

    /** Sets a replacement character for NUL characters read in Unicode strings.

        NUL characters should be replaced to prevent problems with string
        handling. The substitution character is reset to BIFF_DEF_NUL_SUBST_CHAR
        automatically, if a new record is started using the function
        startNextRecord().

        @param cNulSubst
            The character to use for NUL replacement. It is possible to specify
            NUL here. in this case strings are terminated when the first NUL
            occurs during string import.
     */
    void                setNulSubstChar( sal_Unicode cNulSubst = BIFF_DEF_NUL_SUBST_CHAR );

    /** Reads extended unicode string header.

        Detects 8/16-bit mode and all extended info, and seeks to begin of
        the character array.

        @return  Total size of extended string data (formatting and phonetic).
     */
    sal_uInt32          readExtendedUniStringHeader(
                            bool& rb16Bit, bool& rbFonts, bool& rbPhonetic,
                            sal_uInt16& rnFontCount, sal_uInt32& rnPhoneticSize,
                            sal_uInt8 nFlags );

    /** Reads extended unicode string header.

        Detects 8/16-bit mode and seeks to begin of the character array.

        @return  Total size of extended data.
     */
    sal_uInt32          readExtendedUniStringHeader(
                            bool& rb16Bit, sal_uInt8 nFlags );

    /** Reads nChars characters of a BIFF8 string, and returns the string. */
    ::rtl::OUString     readRawUniString( sal_uInt16 nChars, bool b16Bit );
    /** Reads extended header, nChar characters, extended data of a BIFF8
        string, and returns the string. */
    ::rtl::OUString     readUniString( sal_uInt16 nChars, sal_uInt8 nFlags );
    /** Reads 8 bit flags, extended header, nChar characters, extended data of
        a BIFF8 string, and returns the string. */
    ::rtl::OUString     readUniString( sal_uInt16 nChars );
    /** Reads 16 bit character count, 8 bit flags, extended header, character
        array, extended data of a BIFF8 string, and returns the string. */
    ::rtl::OUString     readUniString();

    /** Ignores nChars characters of a BIFF8 string. */
    void                ignoreRawUniString( sal_uInt16 nChars, bool b16Bit );
    /** Ignores extended header, nChar characters, extended data of a BIFF8 string. */
    void                ignoreUniString( sal_uInt16 nChars, sal_uInt8 nFlags );
    /** Ignores 8 bit flags, extended header, nChar characters, extended data
        of a BIFF8 string. */
    void                ignoreUniString( sal_uInt16 nChars );
    /** Ignores 16 bit character count, 8 bit flags, extended header, character
        array, extended data of a BIFF8 string. */
    void                ignoreUniString();

private:
    ::std::auto_ptr< BiffInputStreamImpl > mxImpl;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

