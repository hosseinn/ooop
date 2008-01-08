/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffinputstream.cxx,v $
 *
 *  $Revision: 1.1.2.18 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/23 14:16:15 $
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

#include "oox/xls/biffinputstream.hxx"
#include "oox/core/binaryinputstream.hxx"

using ::rtl::OUString;
using ::rtl::OString;
using ::rtl::OStringToOUString;
using ::oox::core::BinaryInputStream;

namespace oox {
namespace xls {

// ============================================================================

namespace {

/** Buffers the contents of a raw record and encapsulates stream decoding. */
class BiffInputRecordBuffer
{
public:
    explicit            BiffInputRecordBuffer( BinaryInputStream& rInStrm );

    /** Returns the core stream object. */
    inline const BinaryInputStream& getCoreStream() const { return mrInStrm; }

    /** Sets a decoder object and decrypts buffered record data. */
    void                setDecoder( BiffDecoderRef xDecoder );
    /** Returns the current decoder object. */
    inline BiffDecoderRef getDecoder() const { return mxDecoder; }
    /** Enables/disables usage of current decoder. */
    void                enableDecoder( bool bEnable );

    /** Restarts the stream at the passed position. Buffer is invalid until the
        next call of startRecord() or startNextRecord(). */
    void                restartAt( sal_Int64 nPos );

    /** Reads the record header at the passed position. */
    bool                startRecord( sal_Int64 nHeaderPos );
    /** Reads the next record header from the stream. */
    bool                startNextRecord();
    /** Returns the start position of the record header in the core stream. */
    sal_uInt16          getNextRecId();

    /** Returns the start position of the record header in the core stream. */
    inline sal_Int64    getRecHeaderPos() const { return mnHeaderPos; }
    /** Returns the current record identifier. */
    inline sal_uInt16   getRecId() const { return mnRecId; }
    /** Returns the current record size. */
    inline sal_uInt16   getRecSize() const { return mnRecSize; }
    /** Returns the current read position in the current record body. */
    inline sal_uInt16   getRecPos() const { return mnRecPos; }
    /** Returns the number of remaining bytes in the current record body. */
    inline sal_uInt16   getRecLeft() const { return mnRecSize - mnRecPos; }

    /** Reads nBytes bytes to the existing buffer pBuffer. Must NOT overread the buffer. */
    void                read( void* pBuffer, sal_uInt16 nBytes );
    /** Reads a value from the data. Must NOT overread the buffer. */
    template< typename Type >
    inline void         readValue( Type& rnValue );
    /** Ignores nBytes bytes. Must NOT overread the buffer. */
    inline void         skip( sal_uInt16 nBytes );

private:
    /** Updates data buffer from stream, if needed. */
    void                updateBuffer();
    /** Updates decoded data from original data. */
    void                updateDecoded();

private:
    typedef ::std::vector< sal_uInt8 > DataBuffer;

    BinaryInputStream&  mrInStrm;           /// Core input stream.
    DataBuffer          maOriginalData;     /// Original data read from stream.
    DataBuffer          maDecodedData;      /// Decoded data.
    DataBuffer*         mpCurrentData;      /// Points to data buffer currently in use.
    BiffDecoderRef      mxDecoder;          /// Decoder object.
    sal_Int64           mnHeaderPos;        /// Stream start position of current record header.
    sal_Int64           mnBodyPos;          /// Stream start position of current record body.
    sal_Int64           mnBufferBodyPos;    /// Stream start position of buffered data.
    sal_Int64           mnNextHeaderPos;    /// Stream start position of next record header.
    sal_uInt16          mnRecId;            /// Current record identifier.
    sal_uInt16          mnRecSize;          /// Current record size.
    sal_uInt16          mnRecPos;           /// Current position in record body.
    bool                mbValidHeader;      /// True = valid record header.
};

// ----------------------------------------------------------------------------

BiffInputRecordBuffer::BiffInputRecordBuffer( BinaryInputStream& rInStrm ) :
    mrInStrm( rInStrm ),
    mpCurrentData( 0 ),
    mnHeaderPos( -1 ),
    mnBodyPos( 0 ),
    mnBufferBodyPos( 0 ),
    mnNextHeaderPos( 0 ),
    mnRecId( BIFF_ID_UNKNOWN ),
    mnRecSize( 0 ),
    mnRecPos( 0 ),
    mbValidHeader( false )
{
    OSL_ENSURE( mrInStrm.isSeekable(), "BiffInputRecordBuffer::BiffInputRecordBuffer - stream must be seekable" );
    mrInStrm.seek( 0 );
    maOriginalData.reserve( SAL_MAX_UINT16 );
    maDecodedData.reserve( SAL_MAX_UINT16 );
    enableDecoder( false );     // updates mpCurrentData
}

void BiffInputRecordBuffer::restartAt( sal_Int64 nPos )
{
    mnHeaderPos = -1;
    mnBodyPos = mnBufferBodyPos = 0;
    mnNextHeaderPos = nPos;
    mnRecId = BIFF_ID_UNKNOWN;
    mnRecSize = mnRecPos = 0;
    mbValidHeader = false;
}

void BiffInputRecordBuffer::setDecoder( BiffDecoderRef xDecoder )
{
    mxDecoder = xDecoder;
    enableDecoder( true );
    updateDecoded();
}

void BiffInputRecordBuffer::enableDecoder( bool bEnable )
{
    mpCurrentData = (bEnable && mxDecoder.get() && mxDecoder->isValid()) ? &maDecodedData : &maOriginalData;
}

bool BiffInputRecordBuffer::startRecord( sal_Int64 nHeaderPos )
{
    mbValidHeader = (0 <= nHeaderPos) && (nHeaderPos + 4 <= mrInStrm.getLength());
    if( mbValidHeader )
    {
        mnHeaderPos = nHeaderPos;
        mrInStrm.seek( nHeaderPos );
        mrInStrm >> mnRecId >> mnRecSize;
        mnBodyPos = mrInStrm.tell();
        mnNextHeaderPos = mnBodyPos + mnRecSize;
        mbValidHeader = mnNextHeaderPos <= mrInStrm.getLength();
    }
    if( !mbValidHeader )
    {
        mnHeaderPos = mnBodyPos = -1;
        mnNextHeaderPos = 0;
        mnRecId = BIFF_ID_UNKNOWN;
        mnRecSize = 0;
    }
    mnRecPos = 0;
    return mbValidHeader;
}

bool BiffInputRecordBuffer::startNextRecord()
{
    return startRecord( mnNextHeaderPos );
}

sal_uInt16 BiffInputRecordBuffer::getNextRecId()
{
    sal_uInt16 nRecId = BIFF_ID_UNKNOWN;
    if( mbValidHeader && (mnNextHeaderPos + 4 <= mrInStrm.getLength()) )
    {
        mrInStrm.seek( mnNextHeaderPos );
        mrInStrm >> nRecId;
    }
    return nRecId;
}

void BiffInputRecordBuffer::read( void* pBuffer, sal_uInt16 nBytes )
{
    updateBuffer();
    OSL_ENSURE( nBytes > 0, "BiffInputRecordBuffer::read - nothing to read" );
    OSL_ENSURE( mnRecSize - mnRecPos >= nBytes, "BiffInputRecordBuffer::read - buffer overflow" );
    memcpy( pBuffer, &(*mpCurrentData)[ mnRecPos ], nBytes );
    mnRecPos = mnRecPos + nBytes;
}

template< typename Type >
inline void BiffInputRecordBuffer::readValue( Type& rnValue )
{
    read( &rnValue, static_cast< sal_uInt16 >( sizeof( Type ) ) );
    ::oox::core::ByteOrderConverter::convertLittleEndian( rnValue );
}

inline void BiffInputRecordBuffer::skip( sal_uInt16 nBytes )
{
    OSL_ENSURE( mnRecSize - mnRecPos >= nBytes, "BiffInputRecordBuffer::skip - buffer overflow" );
    mnRecPos = mnRecPos + nBytes;
}

void BiffInputRecordBuffer::updateBuffer()
{
    OSL_ENSURE( mbValidHeader, "BiffInputRecordBuffer::updateBuffer - invalid access" );
    if( mnBodyPos != mnBufferBodyPos )
    {
        mrInStrm.seek( mnBodyPos );
        maOriginalData.resize( mnRecSize );
        if( mnRecSize > 0 )
            mrInStrm.read( &maOriginalData.front(), static_cast< sal_Int32 >( mnRecSize ) );
        mnBufferBodyPos = mnBodyPos;
        updateDecoded();
    }
}

void BiffInputRecordBuffer::updateDecoded()
{
    if( mxDecoder.get() && mxDecoder->isValid() )
    {
        maDecodedData.resize( mnRecSize );
        if( mnRecSize > 0 )
            mxDecoder->decode( &maDecodedData.front(), &maOriginalData.front(), mnBodyPos, mnRecSize );
    }
}

} // namespace

// ============================================================================

/** Internal implementation of the BiffInputStream class. */
class BiffInputStreamImpl
{
public:
    explicit            BiffInputStreamImpl(
                            BinaryInputStream& rInStream,
                            bool bContLookup = true );

    // record control ---------------------------------------------------------

    /** Sets stream pointer to the start of the next record content. */
    bool                startNextRecord();
    /** Sets stream pointer to begin of record content. */
    void                resetRecord( bool bContLookup, sal_uInt16 nAltContId );
    /** Sets stream pointer before specified record and invalidates stream. */
    void                rewindToRecord( sal_Int64 nRecHandle );

    // decoder ----------------------------------------------------------------

    /** Sets a new decoder object. */
    void                setDecoder( BiffDecoderRef xDecoder );
    /** Returns the current decoder object. */
    BiffDecoderRef      getDecoder() const;
    /** Enables/disables usage of current decoder. */
    void                enableDecoder( bool bEnable );

    // stream/record state and info -------------------------------------------

    /** Returns record reading state: false = record overread. */
    inline bool         isValid() const { return mbValid; }
    /** Returns the current record identifier. */
    inline sal_uInt16   getRecId() const { return mnRecId; }
    /** Returns the position inside of the whole record content. */
    sal_uInt32          getRecPos() const;
    /** Returns the data size of the whole record without record headers. */
    sal_uInt32          getRecSize();
    /** Returns remaining data size of the whole record without record headers. */
    sal_uInt32          getRecLeft();
    /** Returns a unique handle for the current record. */
    inline sal_Int64    getRecHandle() const { return mnRecHandle; }
    /** Returns the record identifier of the following record. */
    sal_uInt16          getNextRecId();

    /** Returns the absolute core stream position. */
    sal_Int64           getCoreStreamPos() const;
    /** Returns the stream size. */
    sal_Int64           getCoreStreamSize() const;

    // stream read access -----------------------------------------------------

    /** Reads nBytes bytes to the existing(!) buffer pData. */
    sal_uInt32          read( void* pData, sal_uInt32 nBytes );
    /** Reads a value from the stream and converts it to platform byte order. */
    template< typename Type >
    void                readValue( Type& rnValue );

    // seeking ----------------------------------------------------------------

    /** Seeks absolute in record content to the specified position. */
    void                seek( sal_uInt32 nRecPos );
    /** Seeks forward inside the current record. */
    void                ignore( sal_uInt32 nBytes );

    // character arrays -------------------------------------------------------

    /** Reads nChar byte characters and returns the string. */
    OString             readCharArray( sal_uInt16 nChars );
    /** Reads nChars Unicode characters and returns the string. */
    ::rtl::OUString     readUnicodeArray( sal_uInt16 nChars );

    // Unicode strings --------------------------------------------------------

    /** Sets a replacement character for NUL characters read in Unicode strings. */
    inline void         setNulSubstChar( sal_Unicode cNulSubst ) { mcNulSubst = cNulSubst; }
    /** Reads extended unicode string header. */
    sal_uInt32          readExtendedUniStringHeader(
                            bool& rb16Bit, bool& rbFonts, bool& rbPhonetic,
                            sal_uInt16& rnFontCount, sal_uInt32& rnPhoneticSize,
                            sal_uInt8 nFlags );
    /** Reads nChars characters and returns the string. */
    OUString            readRawUniString( sal_uInt16 nChars, bool b16Bit );
    /** Ignores nChars characters. */
    void                ignoreRawUniString( sal_uInt16 nChars, bool b16Bit );

    // private ----------------------------------------------------------------
private:
    /** Initializes all members after base stream has been seeked to new record. */
    void                setupRecord();
    /** Restarts the current record from the beginning. */
    void                restartRecord( bool bInvalidateRecSize );
    /** Returns true, if stream was able to start a valid record. */
    inline bool         isInRecord() const { return mnRecHandle >= 0; }

    /** Returns true, if the passed ID is real or alternative continuation record ID. */
    bool                isContinueId( sal_uInt16 nRecId ) const;
    /** Goes to start of the next CONTINUE record.
        @descr  Stream must be located at the end of a raw record, and handling
        of CONTINUE records must be enabled.
        @return  Copy of mbValid. */
    bool                jumpToNextContinue();
    /** Goes to start of the next CONTINUE record while reading strings.
        @descr  Stream must be located at the end of a raw record. If reading
        has been started in a CONTINUE record, jumps to an existing following
        CONTINUE record, even if handling of CONTINUE records is disabled (this
        is a special handling for TXO string data). Reads additional Unicode
        flag byte at start of the new raw record and sets or resets rb16Bit.
        @return  Copy of mbValid. */
    bool                jumpToNextStringContinue( bool& rb16Bit );

    /** Ensures that reading nBytes bytes is possible with next stream access.
        @descr  Stream must be located at the end of a raw record, and handling
        of CONTINUE records must be enabled.
        @return  Copy of mbValid. */
    bool                ensureRawReadSize( sal_uInt16 nBytes );
    /** Returns the maximum size of raw data possible to read in one block. */
    sal_uInt16          getMaxRawReadSize( sal_uInt32 nBytes ) const;

private:
    BiffInputRecordBuffer maRecBuffer;      /// Raw record data buffer.

    sal_uInt32          mnCurrRecSize;      /// Helper for record size and position.
    sal_uInt32          mnComplRecSize;     /// Size of complete record data (with CONTINUEs).
    bool                mbHasComplRec;      /// True = mnComplRecSize is valid.

    sal_Int64           mnRecHandle;        /// Handle of current record.
    sal_uInt16          mnRecId;            /// Identifier of current record (not the CONTINUE ID).
    sal_uInt16          mnAltContId;        /// Alternative identifier for content continuation records.

    sal_Unicode         mcNulSubst;         /// Replacement for NUL characters.

    bool                mbCont;             /// True = automatic CONTINUE lookup enabled.
    bool                mbValid;            /// True = last stream operation successful (no overread).
};

// ----------------------------------------------------------------------------

BiffInputStreamImpl::BiffInputStreamImpl( BinaryInputStream& rInStream, bool bContLookup ) :
    maRecBuffer( rInStream ),
    mnCurrRecSize( 0 ),
    mnComplRecSize( 0 ),
    mbHasComplRec( false ),
    mnRecHandle( -1 ),
    mnRecId( BIFF_ID_UNKNOWN ),
    mnAltContId( BIFF_ID_UNKNOWN ),
    mcNulSubst( BIFF_DEF_NUL_SUBST_CHAR ),
    mbCont( bContLookup ),
    mbValid( false )
{
}

// record control -------------------------------------------------------------

bool BiffInputStreamImpl::startNextRecord()
{
    bool bValidRec = false;
    /*  #i4266# ignore zero records (id==len==0) (e.g. the application
        "Crystal Report" writes zero records between other records) */
    bool bIsZeroRec = false;
    do
    {
        // record header is never encrypted
        maRecBuffer.enableDecoder( false );
        // read header of next raw record, returns false at end of stream
        bValidRec = maRecBuffer.startNextRecord();
        // ignore record, if identifier and size are zero
        bIsZeroRec = (maRecBuffer.getRecId() == 0) && (maRecBuffer.getRecSize() == 0);
    }
    while( bValidRec && ((mbCont && isContinueId( maRecBuffer.getRecId() )) || bIsZeroRec) );

    // setup other class members
    setupRecord();
    return isInRecord();
}

void BiffInputStreamImpl::resetRecord( bool bContLookup, sal_uInt16 nAltContId )
{
    if( isInRecord() )
    {
        mbCont = bContLookup;
        mnAltContId = nAltContId;
        restartRecord( true );
        maRecBuffer.enableDecoder( true );
    }
}

void BiffInputStreamImpl::rewindToRecord( sal_Int64 nRecHandle )
{
    if( nRecHandle >= 0 )
    {
        maRecBuffer.restartAt( nRecHandle );
        mnRecHandle = -1;
        mbValid = false;
    }
}

// decoder --------------------------------------------------------------------

void BiffInputStreamImpl::setDecoder( BiffDecoderRef xDecoder )
{
    maRecBuffer.setDecoder( xDecoder );
}

BiffDecoderRef BiffInputStreamImpl::getDecoder() const
{
    return maRecBuffer.getDecoder();
}

void BiffInputStreamImpl::enableDecoder( bool bEnable )
{
    maRecBuffer.enableDecoder( bEnable );
}

// stream/record state and info -----------------------------------------------

sal_uInt32 BiffInputStreamImpl::getRecPos() const
{
    return mbValid ? (mnCurrRecSize - maRecBuffer.getRecLeft()) : BIFF_REC_SEEK_TO_END;
}

sal_uInt32 BiffInputStreamImpl::getRecSize()
{
    if( !mbHasComplRec )
    {
        sal_uInt32 nCurrPos = getRecPos();      // save current position in record
        while( jumpToNextContinue() );          // jumpToNextContinue() adds up mnCurrRecSize
        mnComplRecSize = mnCurrRecSize;
        mbHasComplRec = true;
        seek( nCurrPos );                       // restore position, seek() resets old mbValid state
    }
    return mnComplRecSize;
}

sal_uInt32 BiffInputStreamImpl::getRecLeft()
{
    return mbValid ? (getRecSize() - getRecPos()) : 0;
}

sal_uInt16 BiffInputStreamImpl::getNextRecId()
{
    sal_uInt16 nRecId = BIFF_ID_UNKNOWN;
    if( isInRecord() )
    {
        sal_uInt32 nCurrPos = getRecPos();      // save current position in record
        while( jumpToNextContinue() );          // skip following CONTINUE records
        if( maRecBuffer.startNextRecord() )     // read header of next record
            nRecId = maRecBuffer.getRecId();
        seek( nCurrPos );                       // restore position, seek() resets old mbValid state
    }
    return nRecId;
}

sal_Int64 BiffInputStreamImpl::getCoreStreamPos() const
{
    return maRecBuffer.getCoreStream().tell();
}

sal_Int64 BiffInputStreamImpl::getCoreStreamSize() const
{
    return maRecBuffer.getCoreStream().getLength();
}

// stream read access ---------------------------------------------------------

sal_uInt32 BiffInputStreamImpl::read( void* pData, sal_uInt32 nBytes )
{
    sal_uInt32 nRet = 0;
    if( mbValid && pData && (nBytes > 0) )
    {
        sal_uInt8* pnBuffer = reinterpret_cast< sal_uInt8* >( pData );
        sal_uInt32 nBytesLeft = nBytes;

        while( mbValid && (nBytesLeft > 0) )
        {
            sal_uInt16 nReadSize = getMaxRawReadSize( nBytesLeft );
            // check nReadSize, stream may already be located at end of a raw record
            if( nReadSize > 0 )
            {
                maRecBuffer.read( pnBuffer, nReadSize );
                nRet += nReadSize;
                pnBuffer += nReadSize;
                nBytesLeft -= nReadSize;
            }
            if( nBytesLeft > 0 )
                jumpToNextContinue();
            OSL_ENSURE( mbValid, "BiffInputStreamImpl::read - record overread" );
        }
    }
    return nRet;
}

template< typename Type >
void BiffInputStreamImpl::readValue( Type& rnValue )
{
    if( ensureRawReadSize( static_cast< sal_uInt16 >( sizeof( Type ) ) ) )
        maRecBuffer.readValue( rnValue );
}

// seeking --------------------------------------------------------------------

void BiffInputStreamImpl::seek( sal_uInt32 nRecPos )
{
    if( isInRecord() )
    {
        if( !mbValid || (nRecPos < getRecPos()) )
            restartRecord( false );
        if( mbValid && (nRecPos > getRecPos()) )
            ignore( nRecPos - getRecPos() );
    }
}

void BiffInputStreamImpl::ignore( sal_uInt32 nBytes )
{
    sal_uInt32 nBytesLeft = nBytes;
    while( mbValid && (nBytesLeft > 0) )
    {
        sal_uInt16 nSkipSize = getMaxRawReadSize( nBytesLeft );
        maRecBuffer.skip( nSkipSize );
        nBytesLeft -= nSkipSize;
        if( nBytesLeft > 0 )
            jumpToNextContinue();
        OSL_ENSURE( mbValid, "BiffInputStreamImpl::ignore - record overread" );
    }
}

// character arrays -----------------------------------------------------------

OString BiffInputStreamImpl::readCharArray( sal_uInt16 nChars )
{
    ::std::vector< sal_Char > aBuffer( static_cast< size_t >( nChars ) + 1 );
    size_t nCharsRead = static_cast< size_t >( read( &aBuffer.front(), nChars ) );
    aBuffer[ nCharsRead ] = 0;
    return OString( &aBuffer.front() );
}

OUString BiffInputStreamImpl::readUnicodeArray( sal_uInt16 nChars )
{
    ::std::vector< sal_Unicode > aBuffer;
    aBuffer.reserve( static_cast< size_t >( nChars ) + 1 );
    for( sal_uInt16 nCharIdx = 0; mbValid && (nCharIdx < nChars); ++nCharIdx )
    {
        sal_uInt16 nChar;
        readValue( nChar );
        aBuffer.push_back( static_cast< sal_Unicode >( nChar ) );
    }
    aBuffer.push_back( 0 );
    return OUString( &aBuffer.front() );
}

// Unicode strings ------------------------------------------------------------

/*  Unicode string import:

    - Look for CONTINUE records even if CONTINUE handling disabled
      (only if inside of a CONTINUE record - for TXO import).
    - No overread assertions (for Applix wrong string length export bug).

    Structure of an Excel unicode string:
    (1) 2 byte character count
    (2) 1 byte flags (16-bit-characters, rich string, far-east string)
    (3) if rich-string flag set: 2 byte number of font ids
    (4) if far-east flag set: 4 byte far east data size
    (5) character array
    (6) if rich-string flag set: 4 * (number of font ids) byte
    (7) if far-east flag set: (far east data size) byte
    header = (1), (2)
    extended header = (3), (4)
    extended data = (6), (7)
 */

sal_uInt32 BiffInputStreamImpl::readExtendedUniStringHeader(
        bool& rb16Bit, bool& rbFonts, bool& rbPhonetic,
        sal_uInt16& rnFontCount, sal_uInt32& rnPhoneticSize, sal_uInt8 nFlags )
{
    OSL_ENSURE( !getFlag( nFlags, BIFF_STRF_UNKNOWN ), "BiffInputStreamImpl::readExtendedUniStringHeader - unknown flags" );
    rb16Bit = getFlag( nFlags, BIFF_STRF_16BIT );
    rbFonts = getFlag( nFlags, BIFF_STRF_RICH );
    rbPhonetic = getFlag( nFlags, BIFF_STRF_PHONETIC );
    rnFontCount = 0;
    if( rbFonts ) readValue( rnFontCount );
    rnPhoneticSize = 0;
    if( rbPhonetic ) readValue( rnPhoneticSize );
    return rnPhoneticSize + 4 * rnFontCount;
}

OUString BiffInputStreamImpl::readRawUniString( sal_uInt16 nChars, bool b16Bit )
{
    ::std::vector< sal_Unicode > aCharVec;
    aCharVec.reserve( nChars + 1 );

    sal_uInt16 nCharsLeft = nChars;
    while( isValid() && (nCharsLeft > 0) )
    {
        sal_uInt16 nPortionCount = 0;
        if( b16Bit )
        {
            nPortionCount = ::std::min< sal_uInt16 >( nCharsLeft, maRecBuffer.getRecLeft() / 2 );
            OSL_ENSURE( (nPortionCount <= nCharsLeft) || ((maRecBuffer.getRecLeft() & 1) == 0),
                "BiffInputStreamImpl::readRawUniString - missing a byte" );
            // read the character array
            sal_uInt16 nReadChar;
            for( sal_uInt16 nCharIdx = 0; isValid() && (nCharIdx < nPortionCount); ++nCharIdx )
            {
                readValue( nReadChar );
                aCharVec.push_back( (nReadChar == 0) ? mcNulSubst : static_cast< sal_Unicode >( nReadChar ) );
            }
        }
        else
        {
            nPortionCount = getMaxRawReadSize( nCharsLeft );
            // read the character array
            sal_uInt8 nReadChar;
            for( sal_uInt16 nCharIdx = 0; isValid() && (nCharIdx < nPortionCount); ++nCharIdx )
            {
                readValue( nReadChar );
                aCharVec.push_back( (nReadChar == 0) ? mcNulSubst : static_cast< sal_Unicode >( nReadChar ) );
            }
        }

        // prepare for next CONTINUE record
        nCharsLeft = nCharsLeft - nPortionCount;
        if( nCharsLeft > 0 )
            jumpToNextStringContinue( b16Bit );
    }

    // string may contain embedded NUL characters, do not create the OUString by length of vector
    aCharVec.push_back( 0 );
    return OUString( &aCharVec.front() );
}

void BiffInputStreamImpl::ignoreRawUniString( sal_uInt16 nChars, bool b16Bit )
{
    sal_uInt16 nCharsLeft = nChars;

    while( isValid() && (nCharsLeft > 0) )
    {
        sal_uInt16 nPortionCount;
        if( b16Bit )
        {
            nPortionCount = ::std::min< sal_uInt16 >( nCharsLeft, maRecBuffer.getRecLeft() / 2 );
            OSL_ENSURE( (nPortionCount <= nCharsLeft) || ((maRecBuffer.getRecLeft() & 1) == 0),
                "BiffInputStreamImpl::ignoreRawUniString - missing a byte" );
            ignore( 2 * nPortionCount );
        }
        else
        {
            nPortionCount = getMaxRawReadSize( nCharsLeft );
            ignore( nPortionCount );
        }

        // prepare for next CONTINUE record
        nCharsLeft = nCharsLeft - nPortionCount;
        if( nCharsLeft > 0 )
            jumpToNextStringContinue( b16Bit );
    }
}

// private --------------------------------------------------------------------

void BiffInputStreamImpl::setupRecord()
{
    // initialize class members
    mnRecHandle = maRecBuffer.getRecHeaderPos();
    mnRecId = maRecBuffer.getRecId();
    mnAltContId = BIFF_ID_UNKNOWN;
    mnCurrRecSize = mnComplRecSize = maRecBuffer.getRecSize();
    mbHasComplRec = !mbCont;
    mbValid = isInRecord();
    setNulSubstChar( BIFF_DEF_NUL_SUBST_CHAR );
    // enable decoder in new record
    enableDecoder( true );
}

void BiffInputStreamImpl::restartRecord( bool bInvalidateRecSize )
{
    if( isInRecord() )
    {
        maRecBuffer.startRecord( getRecHandle() );
        mnCurrRecSize = maRecBuffer.getRecSize();
        if( bInvalidateRecSize )
        {
            mnComplRecSize = mnCurrRecSize;
            mbHasComplRec = !mbCont;
        }
        mbValid = true;
    }
}

bool BiffInputStreamImpl::isContinueId( sal_uInt16 nRecId ) const
{
    return (nRecId == BIFF_ID_CONT) || (nRecId == mnAltContId);
}

bool BiffInputStreamImpl::jumpToNextContinue()
{
    mbValid = mbValid && mbCont && isContinueId( maRecBuffer.getNextRecId() ) && maRecBuffer.startNextRecord();
    if( mbValid )
        mnCurrRecSize += maRecBuffer.getRecSize();
    return mbValid;
}

bool BiffInputStreamImpl::jumpToNextStringContinue( bool& rb16Bit )
{
    OSL_ENSURE( maRecBuffer.getRecLeft() == 0, "BiffInputStreamImpl::jumpToNextStringContinue - unexpected garbage" );

    if( mbCont && (getRecLeft() > 0) )
    {
        jumpToNextContinue();
    }
    else if( mnRecId == BIFF_ID_CONT )
    {
        /*  CONTINUE handling is off, but we have started reading in a CONTINUE
            record -> start next CONTINUE for TXO import. We really start a new
            record here - no chance to return to string origin. */
        mbValid = mbValid && (maRecBuffer.getNextRecId() == BIFF_ID_CONT) && maRecBuffer.startNextRecord();
        if( mbValid )
            setupRecord();
    }

    // trying to read the flags invalidates stream, if no CONTINUE record has been found
    sal_uInt8 nFlags;
    readValue( nFlags );
    rb16Bit = getFlag( nFlags, BIFF_STRF_16BIT );
    return mbValid;
}

bool BiffInputStreamImpl::ensureRawReadSize( sal_uInt16 nBytes )
{
    if( mbValid && (nBytes > 0) )
    {
        while( mbValid && (maRecBuffer.getRecLeft() == 0) ) jumpToNextContinue();
        mbValid = mbValid && (nBytes <= maRecBuffer.getRecLeft());
        OSL_ENSURE( mbValid, "BiffInputStreamImpl::ensureRawReadSize - record overread" );
    }
    return mbValid;
}

sal_uInt16 BiffInputStreamImpl::getMaxRawReadSize( sal_uInt32 nBytes ) const
{
    return static_cast< sal_uInt16 >( ::std::min< sal_uInt32 >( nBytes, maRecBuffer.getRecLeft() ) );
}

// ============================================================================
// ============================================================================

BiffInputStream::BiffInputStream( BinaryInputStream& rInStream, bool bContLookup ) :
    mxImpl( new BiffInputStreamImpl( rInStream, bContLookup ) )
{
}

BiffInputStream::~BiffInputStream()
{
}

// record control -------------------------------------------------------------

bool BiffInputStream::startNextRecord()
{
    return mxImpl->startNextRecord();
}

bool BiffInputStream::startRecordByHandle( sal_Int64 nRecHandle )
{
    mxImpl->rewindToRecord( nRecHandle );
    return mxImpl->startNextRecord();
}

void BiffInputStream::resetRecord( bool bContLookup, sal_uInt16 nAltContId )
{
    mxImpl->resetRecord( bContLookup, nAltContId );
}

void BiffInputStream::rewindRecord()
{
    mxImpl->rewindToRecord( getRecHandle() );
}

// decoder --------------------------------------------------------------------

void BiffInputStream::setDecoder( BiffDecoderRef xDecoder )
{
    mxImpl->setDecoder( xDecoder );
}

BiffDecoderRef BiffInputStream::getDecoder() const
{
    return mxImpl->getDecoder();
}

void BiffInputStream::enableDecoder( bool bEnable )
{
    mxImpl->enableDecoder( bEnable );
}

// stream/record state and info -----------------------------------------------

bool BiffInputStream::isValid() const
{
    return mxImpl->isValid();
}

sal_uInt16 BiffInputStream::getRecId() const
{
    return mxImpl->getRecId();
}

sal_uInt32 BiffInputStream::getRecPos() const
{
    return mxImpl->getRecPos();
}

sal_uInt32 BiffInputStream::getRecSize() const
{
    return mxImpl->getRecSize();
}

sal_uInt32 BiffInputStream::getRecLeft() const
{
    return mxImpl->getRecLeft();
}

sal_Int64 BiffInputStream::getRecHandle() const
{
    return mxImpl->getRecHandle();
}

sal_uInt16 BiffInputStream::getNextRecId() const
{
    return mxImpl->getNextRecId();
}

sal_Int64 BiffInputStream::getCoreStreamPos() const
{
    return mxImpl->getCoreStreamPos();
}

sal_Int64 BiffInputStream::getCoreStreamSize() const
{
    return mxImpl->getCoreStreamSize();
}

// stream read access ---------------------------------------------------------

sal_uInt32 BiffInputStream::read( void* pData, sal_uInt32 nBytes )
{
    return mxImpl->read( pData, nBytes );
}

BiffInputStream& BiffInputStream::operator>>( sal_Int8& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_uInt8& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_Int16& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_uInt16& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_Int32& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_uInt32& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_Int64& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( sal_uInt64& rnValue )
{
    mxImpl->readValue( rnValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( float& rfValue )
{
    mxImpl->readValue( rfValue );
    return *this;
}

BiffInputStream& BiffInputStream::operator>>( double& rfValue )
{
    mxImpl->readValue( rfValue );
    return *this;
}

sal_Int8 BiffInputStream::readInt8()
{
    sal_Int8 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_uInt8 BiffInputStream::readuInt8()
{
    sal_uInt8 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_Int16 BiffInputStream::readInt16()
{
    sal_Int16 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_uInt16 BiffInputStream::readuInt16()
{
    sal_uInt16 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_Int32 BiffInputStream::readInt32()
{
    sal_Int32 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_uInt32 BiffInputStream::readuInt32()
{
    sal_uInt32 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_Int64 BiffInputStream::readInt64()
{
    sal_Int64 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

sal_uInt64 BiffInputStream::readuInt64()
{
    sal_uInt64 nValue;
    mxImpl->readValue( nValue );
    return nValue;
}

float BiffInputStream::readFloat()
{
    float fValue;
    mxImpl->readValue( fValue );
    return fValue;
}

double BiffInputStream::readDouble()
{
    double fValue;
    mxImpl->readValue( fValue );
    return fValue;
}

// seeking --------------------------------------------------------------------

void BiffInputStream::seek( sal_uInt32 nRecPos )
{
    mxImpl->seek( nRecPos );
}

void BiffInputStream::ignore( sal_uInt32 nBytes )
{
    mxImpl->ignore( nBytes );
}

// character arrays -----------------------------------------------------------

OString BiffInputStream::readCharArray( sal_uInt16 nChars )
{
    return mxImpl->readCharArray( nChars );
}

OUString BiffInputStream::readCharArray( sal_uInt16 nChars, rtl_TextEncoding eTextEnc )
{
    return OStringToOUString( readCharArray( nChars ), eTextEnc );
}

OUString BiffInputStream::readUnicodeArray( sal_uInt16 nChars )
{
    return mxImpl->readUnicodeArray( nChars );
}

// byte strings ---------------------------------------------------------------

OString BiffInputStream::readByteString( bool b16BitLen )
{
    sal_uInt16 nStrLen = b16BitLen ? readuInt16() : readuInt8();
    return mxImpl->readCharArray( nStrLen );
}

OUString BiffInputStream::readByteString( bool b16BitLen, rtl_TextEncoding eTextEnc )
{
    return OStringToOUString( readByteString( b16BitLen ), eTextEnc );
}

void BiffInputStream::ignoreByteString( bool b16BitLen )
{
    sal_uInt16 nStrLen = b16BitLen ? readuInt16() : readuInt8();
    mxImpl->ignore( nStrLen );
}

// Unicode strings ------------------------------------------------------------

void BiffInputStream::setNulSubstChar( sal_Unicode cNulSubst )
{
    mxImpl->setNulSubstChar( cNulSubst );
}

sal_uInt32 BiffInputStream::readExtendedUniStringHeader(
        bool& rb16Bit, bool& rbFonts, bool& rbPhonetic,
        sal_uInt16& rnFontCount, sal_uInt32& rnPhoneticSize, sal_uInt8 nFlags )
{
    return mxImpl->readExtendedUniStringHeader( rb16Bit, rbFonts, rbPhonetic, rnFontCount, rnPhoneticSize, nFlags );
}

sal_uInt32 BiffInputStream::readExtendedUniStringHeader( bool& rb16Bit, sal_uInt8 nFlags )
{
    bool bFonts, bPhonetic;
    sal_uInt16 nFontCount;
    sal_uInt32 nPhoneticSize;
    return mxImpl->readExtendedUniStringHeader( rb16Bit, bFonts, bPhonetic, nFontCount, nPhoneticSize, nFlags );
}

OUString BiffInputStream::readRawUniString( sal_uInt16 nChars, bool b16Bit )
{
    return mxImpl->readRawUniString( nChars, b16Bit );
}

OUString BiffInputStream::readUniString( sal_uInt16 nChars, sal_uInt8 nFlags )
{
    bool b16Bit;
    sal_uInt32 nExtSize = readExtendedUniStringHeader( b16Bit, nFlags );
    OUString aStr = mxImpl->readRawUniString( nChars, b16Bit );
    mxImpl->ignore( nExtSize );
    return aStr;
}

OUString BiffInputStream::readUniString( sal_uInt16 nChars )
{
    return readUniString( nChars, readuInt8() );
}

OUString BiffInputStream::readUniString()
{
    return readUniString( readuInt16() );
}

void BiffInputStream::ignoreRawUniString( sal_uInt16 nChars, bool b16Bit )
{
    mxImpl->ignoreRawUniString( nChars, b16Bit );
}

void BiffInputStream::ignoreUniString( sal_uInt16 nChars, sal_uInt8 nFlags )
{
    bool b16Bit;
    sal_uInt32 nExtSize = readExtendedUniStringHeader( b16Bit, nFlags );
    mxImpl->ignoreRawUniString( nChars, b16Bit );
    mxImpl->ignore( nExtSize );
}

void BiffInputStream::ignoreUniString( sal_uInt16 nChars )
{
    ignoreUniString( nChars, readuInt8() );
}

void BiffInputStream::ignoreUniString()
{
    ignoreUniString( readuInt16() );
}

// ============================================================================

} // namespace xls
} // namespace oox

