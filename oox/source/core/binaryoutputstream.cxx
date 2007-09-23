/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: binaryoutputstream.cxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 14:18:29 $
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

#include "oox/core/binaryoutputstream.hxx"
#include <osl/diagnose.h>
#include <com/sun/star/io/XOutputStream.hpp>
#include <com/sun/star/io/XSeekable.hpp>
#include "oox/core/binaryinputstream.hxx"
#include "oox/core/helper.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::io::XOutputStream;

namespace oox {
namespace core {

// ============================================================================

BinaryOutputStream::BinaryOutputStream(
        const Reference< XOutputStream >& rxOutStrm, sal_Int32 nInitialBufferSize ) :
    mxOutStrm( rxOutStrm ),
    mxSeek( rxOutStrm, UNO_QUERY ),
    maBuffer( nInitialBufferSize )
{
}

BinaryOutputStream::~BinaryOutputStream()
{
    if( mxOutStrm.is() )
        mxOutStrm->closeOutput();
}

sal_Int64 BinaryOutputStream::tell() const
{
    try
    {
        return mxSeek.is() ? mxSeek->getPosition() : -1;
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryOutputStream::tell - exception caught" );
    }
    return -1;
}

void BinaryOutputStream::seek( sal_Int64 nPos )
{
    try
    {
        if( mxSeek.is() )
            mxSeek->seek( nPos );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryOutputStream::seek - exception caught" );
    }
}

void BinaryOutputStream::write( const Sequence< sal_Int8 >& rBuffer )
{
    try
    {
        mxOutStrm->writeBytes( rBuffer );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryOutputStream::write - stream read error" );
    }
}

void BinaryOutputStream::write( const void* pBuffer, sal_Int32 nBytes )
{
    if( nBytes > 0 )
    {
        maBuffer.realloc( nBytes );
        memcpy( maBuffer.getArray(), pBuffer, static_cast< size_t >( nBytes ) );
        write( maBuffer );
    }
}

void BinaryOutputStream::copy( BinaryInputStream& rInStrm, sal_Int64 nBytes )
{
    if( rInStrm.is() && (nBytes > 0) )
    {
        sal_Int32 nBufferSize = getLimitedValue< sal_Int32, sal_Int64 >( nBytes, 0, 0x8000 );
        Sequence< sal_Int8 > aBuffer( nBufferSize );
        while( nBytes > 0 )
        {
            sal_Int32 nReadSize = getLimitedValue< sal_Int32, sal_Int64 >( nBytes, 0, nBufferSize );
            rInStrm.read( aBuffer, nReadSize );
            write( aBuffer );
            nBytes -= nReadSize;
        }
    }
}

void BinaryOutputStream::copy( BinaryInputStream& rInStrm )
{
    if( rInStrm.is() && rInStrm.isSeekable() )
        copy( rInStrm, rInStrm.getLength() - rInStrm.tell() );
}

// ============================================================================

} // namespace core
} // namespace oox

