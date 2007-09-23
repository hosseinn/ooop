/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffoutputstream.cxx,v $
 *
 *  $Revision: 1.1.2.1 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/03 14:00:23 $
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

#include "oox/xls/biffoutputstream.hxx"

using ::com::sun::star::uno::Reference;
using ::com::sun::star::io::XOutputStream;

namespace oox {
namespace xls {

// ============================================================================

BiffOutputStream::BiffOutputStream( const Reference< XOutputStream >& /*rxOutStream*/ )
{
}

 BiffOutputStream::~BiffOutputStream()
 {
 }

void BiffOutputStream::setSliceSize( sal_uInt16 /*nSize*/ )
{
}

// stream write access --------------------------------------------------------

sal_uInt32 BiffOutputStream::write( const void* /*pData*/, sal_uInt32 nBytes )
{
    return nBytes;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_Int8 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_uInt8 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_Int16 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_uInt16 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_Int32 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_uInt32 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_Int64 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( sal_uInt64 /*nValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( float /*fValue*/ )
{
    return *this;
}

BiffOutputStream& BiffOutputStream::operator<<( double /*fValue*/ )
{
    return *this;
}

// ============================================================================

} // namespace xls
} // namespace oox

