/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: querytablefragment.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/01 09:58:14 $
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

#include "oox/xls/querytablefragment.hxx"
#include "oox/xls/webquerybuffer.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::xml::sax::SAXException;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

OoxQueryTableFragment::OoxQueryTableFragment(
        const GlobalDataHelper& rGlobalData, const OUString& rFragmentPath ) :
    GlobalFragmentBase( rGlobalData, rFragmentPath )
{
}

bool OoxQueryTableFragment::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XML_ROOT_CONTEXT: return (nElement == XLS_TOKEN( queryTable ));
    }
    return false;
}

void OoxQueryTableFragment::onStartElement( const AttributeList& rAttribs )
{
    switch ( getCurrentContext() )
    {
        case XLS_TOKEN( queryTable ):
            importQueryTable( rAttribs );
        break;
    }
}

void OoxQueryTableFragment::importQueryTable( const AttributeList& rAttribs )
{
    getWebQueries().importQueryTable( rAttribs );
}

} // namespace xls
} // namespace oox
