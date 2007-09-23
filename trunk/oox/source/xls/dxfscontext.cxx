/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: dxfscontext.cxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:20 $
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

#include "oox/xls/dxfscontext.hxx"
#include "oox/xls/excelfragmentbase.hxx"
#include "oox/xls/stylesbuffer.hxx"

using ::rtl::OUString;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

OoxDxfsContext::OoxDxfsContext( const GlobalFragmentBase& rFragment ) :
    GlobalContextBase( rFragment )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxDxfsContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( dxfs ):
            return (nElement == XLS_TOKEN( dxf ));
        case XLS_TOKEN( dxf ):
            return (nElement == XLS_TOKEN( border )) ||
                   (nElement == XLS_TOKEN( font )) ||
                   (nElement == XLS_TOKEN( fill )) ||
                   (nElement == XLS_TOKEN( numFmt )) ||
                   (nElement == XLS_TOKEN( protection ));

        case XLS_TOKEN( border ):
        case XLS_TOKEN( left ):
        case XLS_TOKEN( right ):
        case XLS_TOKEN( top ):
        case XLS_TOKEN( bottom ):
        case XLS_TOKEN( diagonal ):
            return Border::isSupportedContext( nElement, getCurrentContext() );

        case XLS_TOKEN( font ):
        case XLS_TOKEN( color ):
            return Font::isSupportedContext( nElement, getCurrentContext() );

        case XLS_TOKEN( fill ):
        case XLS_TOKEN( patternFill ):
        case XLS_TOKEN( gradientFill ):
        case XLS_TOKEN( fgColor ):
        case XLS_TOKEN( bgColor ):
            return Fill::isSupportedContext( nElement, getCurrentContext() );
    }
    return false;
}

void OoxDxfsContext::onStartElement( const AttributeList& rAttribs )
{
    switch ( getCurrentContext() )
    {
        case XLS_TOKEN( dxf ):
            mpCurDxf = getStyles().importDxf( rAttribs );
        break;

        // shared elements

        case XLS_TOKEN( color ):
            importColor( rAttribs );
        break;

        // font context

        case XLS_TOKEN( font ):
        break;

        // fill context

        case XLS_TOKEN( fill ):
        break;
        case XLS_TOKEN( patternFill ):
            mpCurDxf->getFill()->importPatternFill( rAttribs, false );
        break;
        case XLS_TOKEN( gradientFill ):
            mpCurDxf->getFill()->importGradientFill( rAttribs );
        break;
        case XLS_TOKEN( fgColor ):
            mpCurDxf->getFill()->importFgColor( rAttribs );
        break;
        case XLS_TOKEN( bgColor ):
            mpCurDxf->getFill()->importBgColor( rAttribs );
        break;

        // border context

        case XLS_TOKEN( border ):
            mpCurDxf->getBorder()->importBorder( rAttribs );
        break;
        case XLS_TOKEN( left ):
        case XLS_TOKEN( right ):
        case XLS_TOKEN( top ):
        case XLS_TOKEN( bottom ):
        case XLS_TOKEN( diagonal ):
            mpCurDxf->getBorder()->importStyle( getCurrentContext(), rAttribs );
        break;
    }
}

void OoxDxfsContext::onEndElement( const OUString& /*rChars*/ )
{
    switch ( getCurrentContext() )
    {
        case XLS_TOKEN( dxf ):
            mpCurDxf->getFont()->finalizeImport();
            mpCurDxf->getFill()->finalizeImport();
            mpCurDxf->getBorder()->finalizeImport();
        break;
    }
}

void OoxDxfsContext::importColor( const AttributeList& rAttribs )
{
    switch ( getPreviousContext() )
    {
        case XLS_TOKEN( font ):
            mpCurDxf->getFont()->importAttribs( getCurrentContext(), rAttribs );
        break;
        case XLS_TOKEN( left ):
        case XLS_TOKEN( right ):
        case XLS_TOKEN( top ):
        case XLS_TOKEN( bottom ):
        case XLS_TOKEN( diagonal ):
            mpCurDxf->getBorder()->importColor( getPreviousContext(), rAttribs );
        break;
    }
}

} // namespace xls
} // namespace oox


