/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textliststylecontext.cxx,v $
 *
 *  $Revision: 1.1.2.5 $
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

#include "oox/drawingml/textliststylecontext.hxx"
#include "oox/drawingml/textparagraphpropertiescontext.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"
#include "tokens.hxx"

using ::rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

// --------------------------------------------------------------------

// CT_TextListStyle
TextListStyleContext::TextListStyleContext( const ::oox::core::FragmentHandlerRef& xHandler, oox::drawingml::TextListStylePtr pTextListStylePtr )
: Context( xHandler )
, mpTextListStylePtr( pTextListStylePtr )
{
}

TextListStyleContext::~TextListStyleContext()
{
}

// --------------------------------------------------------------------

void TextListStyleContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
}

// --------------------------------------------------------------------

Reference< XFastContextHandler > TextListStyleContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& rxAttributes ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	TextParagraphPropertiesPtr       pProps;

	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_defPPr:		// CT_TextParagraphProperties
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 0 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_outline1pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getAggregationListStyle()[ 0 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_outline2pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getAggregationListStyle()[ 1 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl1pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 1 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl2pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 2 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl3pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 3 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl4pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 4 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl5pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 5 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl6pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 6 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl7pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 7 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl8pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 8 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
		case NMSP_DRAWINGML|XML_lvl9pPr:
			pProps.reset( new TextParagraphProperties() );
			mpTextListStylePtr->getListStyle()[ 9 ] = pProps;
			xRet.set( new TextParagraphPropertiesContext( this, rxAttributes, pProps ) );
			break;
	}
	if ( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// --------------------------------------------------------------------

} }

