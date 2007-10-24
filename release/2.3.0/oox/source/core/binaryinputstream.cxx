/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: binaryinputstream.cxx,v $
 *
 *  $Revision: 1.1.2.4 $
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

#include "oox/core/binaryinputstream.hxx"
#include <osl/diagnose.h>
#include <com/sun/star/io/XInputStream.hpp>
#include <com/sun/star/io/XSeekable.hpp>

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::io::XInputStream;

namespace oox {
namespace core {

// ============================================================================

BinaryInputStream::BinaryInputStream(
        const Reference< XInputStream >& rxInStrm, sal_Int32 nInitialBufferSize ) :
    mxInStrm( rxInStrm ),
    mxSeek( rxInStrm, UNO_QUERY ),
    maBuffer( nInitialBufferSize )
{
}

BinaryInputStream::~BinaryInputStream()
{
    if( mxInStrm.is() )
        mxInStrm->closeInput();
}

sal_Int64 BinaryInputStream::getLength() const
{
    try
    {
        return mxSeek.is() ? mxSeek->getLength() : -1;
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryInputStream::getLength - exception caught" );
    }
    return -1;
}

sal_Int64 BinaryInputStream::tell() const
{
    try
    {
        return mxSeek.is() ? mxSeek->getPosition() : -1;
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryInputStream::tell - exception caught" );
    }
    return -1;
}

void BinaryInputStream::seek( sal_Int64 nPos )
{
    try
    {
        if( mxSeek.is() )
            mxSeek->seek( nPos );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryInputStream::seek - exception caught" );
    }
}

void BinaryInputStream::skip( sal_Int32 nBytes )
{
    try
    {
        mxInStrm->skipBytes( nBytes );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryInputStream::skip - exception caught" );
    }
}

sal_Int32 BinaryInputStream::read( Sequence< sal_Int8 >& orBuffer, sal_Int32 nBytes )
{
    sal_Int32 nRet = 0;
    try
    {
        nRet = mxInStrm->readBytes( orBuffer, nBytes );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "BinaryInputStream::read - stream read error" );
    }
    return nRet;
}

sal_Int32 BinaryInputStream::read( void* opBuffer, sal_Int32 nBytes )
{
    sal_Int32 nRet = read( maBuffer, nBytes );
    if( nRet > 0 )
        memcpy( opBuffer, maBuffer.getConstArray(), static_cast< size_t >( nRet ) );
    return nRet;
}

// ============================================================================

} // namespace core
} // namespace oox

