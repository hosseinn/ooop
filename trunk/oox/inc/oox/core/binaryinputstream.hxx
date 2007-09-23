/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: binaryinputstream.hxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 14:18:20 $
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

#ifndef OOX_CORE_BINARYINPUTSTREAM_HXX
#define OOX_CORE_BINARYINPUTSTREAM_HXX

#include <com/sun/star/uno/Reference.hxx>
#include <com/sun/star/uno/Sequence.hxx>
#include "oox/core/helper.hxx"

namespace com { namespace sun { namespace star {
    namespace io { class XInputStream; }
    namespace io { class XSeekable; }
} } }

namespace oox {
namespace core {

// ============================================================================

/** Wraps a binary input stream and provides convenient access functions.

    The binary data in the stream is assumed to be in little-endian format.
 */
class BinaryInputStream
{
public:
    explicit            BinaryInputStream(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream >& rxInStrm,
                            sal_Int32 nInitialBufferSize = 0x8000 );

                        ~BinaryInputStream();

    /** Returns true, if the wrapped stream is valid. */
    inline bool         is() const { return mxInStrm.is(); }
    /** Returns true, if the wrapped stream is seekable. */
    inline bool         isSeekable() const { return mxSeek.is(); }
    /** Returns the XInputStream interface of the wrapped input stream. */
    inline ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream >
                        getXInputStream() const { return mxInStrm; }

    /** Returns the size of the stream, if stream is seekable, otherwise -1. */
    sal_Int64           getLength() const;

    /** Returns the current stream position, if stream is seekable, otherwise -1. */
    sal_Int64           tell() const;
    /** Seeks the stream to the passed position, if stream is seekable. */
    void                seek( sal_Int64 nPos );
    /** Seeks the stream forward by the passed number of bytes. This works for
        non-seekable streams too. */
    void                skip( sal_Int32 nBytes );

    /** Reads nBytes bytes to the passed sequence.
        @return  Number of bytes really read. */
    sal_Int32           read( ::com::sun::star::uno::Sequence< sal_Int8 >& orBuffer, sal_Int32 nBytes );
    /** Reads nBytes bytes to the (existing) buffer pBuffer.
        @return  Number of bytes really read. */
    sal_Int32           read( void* opBuffer, sal_Int32 nBytes );

    /** Reads a value from the stream and converts it to platform byte order. */
    template< typename Type >
    void                readValue( Type& ornValue );
    /** Reads a value from the stream and converts it to platform byte order. */
    template< typename Type >
    inline Type         readValue() { Type nValue; readValue( nValue ); return nValue; }

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream >
                        mxInStrm;       /// Reference to the input stream.
    ::com::sun::star::uno::Reference< ::com::sun::star::io::XSeekable >
                        mxSeek;         /// Stream seeking interface.
    ::com::sun::star::uno::Sequence< sal_Int8 >
                        maBuffer;       /// Data buffer for readBytes() calls.
};

// ----------------------------------------------------------------------------

template< typename Type >
void BinaryInputStream::readValue( Type& ornValue )
{
    // can be instanciated for all types supported in class ByteOrderConverter
    read( &ornValue, static_cast< sal_Int32 >( sizeof( Type ) ) );
    ByteOrderConverter::convertLittleEndian( ornValue );
}

template< typename Type >
inline BinaryInputStream& operator>>( BinaryInputStream& rInStrm, Type& ornValue )
{
    rInStrm.readValue( ornValue );
    return rInStrm;
}

// ============================================================================

} // namespace core
} // namespace oox

#endif

