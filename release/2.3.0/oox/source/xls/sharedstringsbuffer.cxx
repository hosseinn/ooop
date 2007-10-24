/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sharedstringsbuffer.cxx,v $
 *
 *  $Revision: 1.1.2.14 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:21 $
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

#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/biffinputstream.hxx"

using ::com::sun::star::uno::Reference;
using ::com::sun::star::text::XText;

namespace oox {
namespace xls {

// ============================================================================

SharedStringsBuffer::SharedStringsBuffer( const GlobalDataHelper& rGlobalData ) :
     GlobalDataHelper( rGlobalData )
{
}

RichStringRef SharedStringsBuffer::createRichString()
{
    RichStringRef xString( new RichString( getGlobalData() ) );
    maStrings.push_back( xString );
    return xString;
}

void SharedStringsBuffer::importSst( BiffInputStream& rStrm )
{
    sal_Int32 nStringCount;
    rStrm.ignore( 4 );
    rStrm >> nStringCount;
    if( nStringCount > 0 )
    {
        maStrings.clear();
        maStrings.reserve( static_cast< size_t >( nStringCount ) );
        for( ; rStrm.isValid() && (nStringCount > 0); --nStringCount )
        {
            RichStringRef xString( new RichString( getGlobalData() ) );
            maStrings.push_back( xString );
            xString->importUniString( rStrm );
        }
    }
}

void SharedStringsBuffer::finalizeImport()
{
    maStrings.forEachMem( &RichString::finalizeImport );
}

void SharedStringsBuffer::convertString( const Reference< XText >& rxText, sal_Int32 nIndex, sal_Int32 nXfId ) const
{
    if( rxText.is() )
        if( const RichString* pString = maStrings.get( nIndex ).get() )
            pString->convert( rxText, nXfId );
}

// ============================================================================

} // namespace xls
} // namespace oox

