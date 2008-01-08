/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: spdefcontext.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: hub $ $Date: 2007/06/19 19:50:46 $
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

#include "oox/drawingml/spdefcontext.hxx"
#include "oox/drawingml/shapepropertiescontext.hxx"
#include "oox/drawingml/textbodypropertiescontext.hxx"
#include "oox/drawingml/textliststylecontext.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

spDefContext::spDefContext( const ::oox::core::FragmentHandlerRef& xHandler, oox::drawingml::ShapePtr pDefaultObjectPtr )
: Context( xHandler )
, mpDefaultObjectPtr( pDefaultObjectPtr )
{
}

Reference< XFastContextHandler > spDefContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_spPr:
		{
			xRet = new ShapePropertiesContext( this, mpDefaultObjectPtr );
			break;
		}
		case NMSP_DRAWINGML|XML_bodyPr:
		{
			xRet = new TextBodyPropertiesContext( this, xAttribs, mpDefaultObjectPtr );
			break;
		}
		case NMSP_DRAWINGML|XML_lstStyle:
			xRet.set( new TextListStyleContext( getHandler(), mpDefaultObjectPtr->getListStyles() ) );
			break;
		case NMSP_DRAWINGML|XML_style:
			break;
	}
	if( !xRet.is() )
		xRet.set( this );

	return xRet;
}

} }
