/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: externallinkfragment.cxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:48 $
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

#include "oox/xls/externallinkfragment.hxx"
#include "oox/xls/externallinkbuffer.hxx"
#include "oox/xls/sheetdatacontext.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::oox::core::AttributeList;
using ::oox::core::Relation;

namespace oox {
namespace xls {

// ============================================================================

OoxExternalLinkFragment::OoxExternalLinkFragment( const GlobalDataHelper& rGlobalData,
        ExternalLink& rExtLink, const OUString& rFragmentPath ) :
    GlobalFragmentBase( rGlobalData, rFragmentPath ),
    mrExtLink( rExtLink ),
    mnCurrSheetId( -1 )
{
}

bool OoxExternalLinkFragment::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XML_ROOT_CONTEXT:
            return  (nElement == XLS_TOKEN( externalLink ));
        case XLS_TOKEN( externalLink ):
            return  (nElement == XLS_TOKEN( externalBook )) ||
                    (nElement == XLS_TOKEN( ddeLink )) ||
                    (nElement == XLS_TOKEN( oleLink ));
        case XLS_TOKEN( externalBook ):
            return  (nElement == XLS_TOKEN( sheetNames )) ||
                    (nElement == XLS_TOKEN( definedNames )) ||
                    (nElement == XLS_TOKEN( sheetDataSet ));
        case XLS_TOKEN( sheetNames ):
            return  (nElement == XLS_TOKEN( sheetName ));
        case XLS_TOKEN( definedNames ):
            return  (nElement == XLS_TOKEN( definedName ));
        case XLS_TOKEN( sheetDataSet ):
            return  (nElement == XLS_TOKEN( sheetData ));
        case XLS_TOKEN( ddeLink ):
            return  (nElement == XLS_TOKEN( ddeItems ));
        case XLS_TOKEN( ddeItems ):
            return  (nElement == XLS_TOKEN( ddeItem ));
        case XLS_TOKEN( oleLink ):
            return  (nElement == XLS_TOKEN( oleItems ));
        case XLS_TOKEN( oleItems ):
            return  (nElement == XLS_TOKEN( oleItem ));
    }
    return false;
}

Reference< XFastContextHandler > OoxExternalLinkFragment::onCreateContext( sal_Int32 nElement, const AttributeList& rAttribs )
{
    switch( nElement )
    {
        case XLS_TOKEN( sheetData ):
        {
            sal_Int32 nSheet = mrExtLink.getSheetIndex( rAttribs.getInteger( XML_sheetId, -1 ) );
            Reference< XSpreadsheet > xSheet = getSheet( nSheet );
            if( xSheet.is() )
                return new OoxExternalSheetDataContext( *this, WorksheetHelper( getGlobalData(), SHEETTYPE_WORKSHEET, xSheet, nSheet ) );
        }
        break;
    }
    return this;
}

void OoxExternalLinkFragment::onStartElement( const AttributeList& rAttribs )
{
    switch ( getCurrentContext() )
    {
        case XLS_TOKEN( externalBook ):
            if( const Relation* pRelation = getRelationById( rAttribs.getString( R_TOKEN( id ) ) ).get() )
                mrExtLink.importExternalBook( rAttribs, pRelation->msTarget );
        break;
        case XLS_TOKEN( sheetName ):
            mrExtLink.importSheetName( rAttribs );
        break;
        case XLS_TOKEN( definedName ):
            mrExtLink.importDefinedName( rAttribs );
        break;
        case XLS_TOKEN( ddeLink ):
            mrExtLink.importDdeLink( rAttribs );
        break;
        case XLS_TOKEN( ddeItem ):
            mrExtLink.importDdeItem( rAttribs );
        break;
        case XLS_TOKEN( oleLink ):
            if( const Relation* pRelation = getRelationById( rAttribs.getString( R_TOKEN( id ) ) ).get() )
                mrExtLink.importOleLink( rAttribs, pRelation->msTarget );
        break;
        case XLS_TOKEN( oleItem ):
            mrExtLink.importOleItem( rAttribs );
        break;
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

