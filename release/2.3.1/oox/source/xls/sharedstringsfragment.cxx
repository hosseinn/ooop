/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sharedstringsfragment.cxx,v $
 *
 *  $Revision: 1.1.2.14 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:49 $
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

#include "oox/xls/sharedstringsfragment.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/richstringcontext.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

// ============================================================================

OoxSharedStringsFragment::OoxSharedStringsFragment(
        const GlobalDataHelper& rGlobalData, const OUString& rFragmentPath ) :
    GlobalFragmentBase( rGlobalData, rFragmentPath )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxSharedStringsFragment::onCanCreateContext( sal_Int32 nElement )
{
    sal_Int32 nCurrContext = getCurrentContext();
    switch( nCurrContext )
    {
        case XML_ROOT_CONTEXT:
            return  (nElement == XLS_TOKEN( sst ));
        case XLS_TOKEN( sst ):
            return  (nElement == XLS_TOKEN( si ));
    }
    return false;
}

Reference< XFastContextHandler > OoxSharedStringsFragment::onCreateContext( sal_Int32 nElement, const AttributeList& /*rAttribs*/ )
{
    switch( nElement )
    {
        case XLS_TOKEN( si ):
            return new OoxRichStringContext( *this, getSharedStrings().createRichString() );
    }
    return this;
}

void OoxSharedStringsFragment::onEndElement( const OUString& )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sst ):
            getSharedStrings().finalizeImport();
        break;
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

