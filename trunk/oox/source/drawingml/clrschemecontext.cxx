/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: clrschemecontext.cxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: sj $ $Date: 2007/07/19 09:49:44 $
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

#include "oox/drawingml/clrschemecontext.hxx"
#include "oox/drawingml/colorchoicecontext.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

clrSchemeContext::clrSchemeContext( const ::oox::core::FragmentHandlerRef& xHandler, const oox::drawingml::ClrSchemePtr pClrSchemePtr )
: Context( xHandler )
, mpClrSchemePtr( pClrSchemePtr )
, mnColor( 0 )
{
	mpClrSchemePtr->setColor( XML_bg1, 0xffffff );	// semantic background and text colors
	mpClrSchemePtr->setColor( XML_bg2, 0xffffff );
	mpClrSchemePtr->setColor( XML_tx1, 0 );
	mpClrSchemePtr->setColor( XML_tx2, 0 );
}

void clrSchemeContext::startFastElement( sal_Int32 /* aElementToken */, const Reference< XFastAttributeList >& /* xAttribs */ ) throw (SAXException, RuntimeException)
{
}

void clrSchemeContext::endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException)
{
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_dk1:
		case NMSP_DRAWINGML|XML_lt1:
		case NMSP_DRAWINGML|XML_dk2:
		case NMSP_DRAWINGML|XML_lt2:
		case NMSP_DRAWINGML|XML_accent1:
		case NMSP_DRAWINGML|XML_accent2:
		case NMSP_DRAWINGML|XML_accent3:
		case NMSP_DRAWINGML|XML_accent4:
		case NMSP_DRAWINGML|XML_accent5:
		case NMSP_DRAWINGML|XML_accent6:
		case NMSP_DRAWINGML|XML_hlink:
		case NMSP_DRAWINGML|XML_folHlink:
		{
			mpClrSchemePtr->setColor( aElementToken & 0xffff, mnColor );
			break;
		}
	}
}

Reference< XFastContextHandler > clrSchemeContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& /* xAttribs */ ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_scrgbClr:	// CT_ScRgbColor
		case NMSP_DRAWINGML|XML_srgbClr:	// CT_SRgbColor
		case NMSP_DRAWINGML|XML_hslClr:	// CT_HslColor
		case NMSP_DRAWINGML|XML_sysClr:	// CT_SystemColor
//		case NMSP_DRAWINGML|XML_schemeClr:	// CT_SchemeColor
		case NMSP_DRAWINGML|XML_prstClr:	// CT_PresetColor
		{
			xRet.set( new colorChoiceContext( getHandler(), mnColor ) );
			break;
		}
	}
	if( !xRet.is() )
		xRet.set( this );

	return xRet;
}

} }
