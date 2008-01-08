/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: binaryoutputstream.hxx,v $
 *
 *  $Revision: 1.1.2.4 $
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

#ifndef OOX_CORE_BINARYOUTPUTSTREAM_HXX
#define OOX_CORE_BINARYOUTPUTSTREAM_HXX

#include <com/sun/star/uno/Reference.hxx>
#include <com/sun/star/uno/Sequence.hxx>
#include "oox/core/helper.hxx"

namespace com { namespace sun { namespace star {
    namespace io { class XOutputStream; }
    namespace io { class XSeekable; }
} } }

namespace oox {
namespace core {

class BinaryInputStream;

// ============================================================================

/** Wraps a binary output stream and provides convenient access functions.

    The binary data in the stream is written in little-endian format.
 */
class BinaryOutputStream
{
public:
    explicit            BinaryOutputStream(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream >& rxOutStrm,
                            sal_Int32 nInitialBufferSize = 0x8000 );

                        ~BinaryOutputStream();

    /** Returns true, if the wrapped stream is valid. */
    inline bool         is() const { return mxOutStrm.is(); }
    /** Returns true, if the wrapped stream is seekable. */
    inline bool         isSeekable() const { return mxSeek.is(); }
    /** Returns the XOutputStream interface of the wrapped output stream. */
    inline ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream >
                        getXOutputStream() const { return mxOutStrm; }

    /** Returns the current stream position, if stream is seekable, otherwise -1. */
    sal_Int64           tell() const;
    /** Seeks the stream to the passed position, if stream is seekable. */
    void                seek( sal_Int64 nPos );

    /** Writes the passed sequence. */
    void                write( const ::com::sun::star::uno::Sequence< sal_Int8 >& rBuffer );
    /** Writes nBytes bytes from the (existing) buffer pBuffer. */
    void                write( const void* pBuffer, sal_Int32 nBytes );

    /** Copies nBytes bytes from the current position of the passed input stream. */
    void                copy( BinaryInputStream& rInStrm, sal_Int64 nBytes );
    /** Copies the passed input stream from its current position to its end. */
    void                copy( BinaryInputStream& rInStrm );

    /** Writes a value to the stream and converts it to platform byte order. */
    template< typename Type >
    void                writeValue( Type nValue );

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream >
                        mxOutStrm;      /// Reference to the output stream.
    ::com::sun::star::uno::Reference< ::com::sun::star::io::XSeekable >
                        mxSeek;         /// Stream seeking interface.
    ::com::sun::star::uno::Sequence< sal_Int8 >
                        maBuffer;       /// Data buffer for readBytes() calls.
};

// ----------------------------------------------------------------------------

template< typename Type >
void BinaryOutputStream::writeValue( Type nValue )
{
    // can be instanciated for all types supported in class ByteOrderConverter
    ByteOrderConverter::convertLittleEndian( nValue );
    write( &nValue, static_cast< sal_Int32 >( sizeof( Type ) ) );
}

// ============================================================================

} // namespace core
} // namespace oox

#endif

