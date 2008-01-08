/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: colorchoicecontext.cxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/02 11:27:15 $
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

#include "oox/drawingml/colorchoicecontext.hxx"
#include "oox/drawingml/clrscheme.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

colorChoiceContext::colorChoiceContext( const ::oox::core::FragmentHandlerRef& xHandler, sal_Int32& rColor )
: Context( xHandler )
, mrColor( rColor )
{
}

void colorChoiceContext::startFastElement( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	switch( aElementToken )
	{
	case NMSP_DRAWINGML|XML_scrgbClr:	// CT_ScRgbColor
	{
		sal_Int32 r = ((xAttribs->getOptionalValue( XML_r ).toInt32() * 256) / 1000) & 0xff;
		sal_Int32 g = ((xAttribs->getOptionalValue( XML_g ).toInt32() * 256) / 1000) & 0xff;
		sal_Int32 b = ((xAttribs->getOptionalValue( XML_b ).toInt32() * 256) / 1000) & 0xff;
		mrColor = (r << 16) | (g << 8) | b;
		break;
	}
	case NMSP_DRAWINGML|XML_srgbClr:	// CT_SRgbColor
	{
		mrColor = xAttribs->getOptionalValue( XML_val ).toInt32( 16 );
		break;
	}
	case NMSP_DRAWINGML|XML_hslClr:	// CT_HslColor
		// todo
		break;
	case NMSP_DRAWINGML|XML_sysClr:	// CT_SystemColor
        if( !ClrScheme::getSystemColor( xAttribs->getOptionalValueToken( XML_val, XML_TOKEN_INVALID ), mrColor ) )
            mrColor = xAttribs->getOptionalValue( XML_lastClr ).toInt32( 16 );
		break;
	case NMSP_DRAWINGML|XML_schemeClr:	// CT_SchemeColor
	{
        mrColor = getHandler()->getFilter()->getSchemeClr( xAttribs->getOptionalValueToken( XML_val, XML_nothing ) );
		break;
	}
	case NMSP_DRAWINGML|XML_prstClr:	// CT_PresetColor
		// todo
		break;

	}
}

Reference< XFastContextHandler > colorChoiceContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	// colorTransformGroup

	// color should be available as rgb in member mnColor already, now modify it depending on
	// the transformation elements

	switch( aElementToken )
	{
	case NMSP_DRAWINGML|XML_tint:		// CT_PositiveFixedPercentage
	case NMSP_DRAWINGML|XML_shade:		// CT_PositiveFixedPercentage
	case NMSP_DRAWINGML|XML_comp:		// CT_ComplementTransform
	case NMSP_DRAWINGML|XML_inv:		// CT_InverseTransform
	case NMSP_DRAWINGML|XML_gray:		// CT_GrayscaleTransform
	case NMSP_DRAWINGML|XML_alpha:		// CT_PositiveFixedPercentage
	case NMSP_DRAWINGML|XML_alphaOff:	// CT_FixedPercentage
	case NMSP_DRAWINGML|XML_alphaMod:	// CT_PositivePercentage
	case NMSP_DRAWINGML|XML_hue:		// CT_PositiveFixedAngle
	case NMSP_DRAWINGML|XML_hueOff:	// CT_Angle
	case NMSP_DRAWINGML|XML_hueMod:	// CT_PositivePercentage
	case NMSP_DRAWINGML|XML_sat:		// CT_Percentage
	case NMSP_DRAWINGML|XML_satOff:	// CT_Percentage
	case NMSP_DRAWINGML|XML_satMod:	// CT_Percentage
	case NMSP_DRAWINGML|XML_lum:		// CT_Percentage
	case NMSP_DRAWINGML|XML_lumOff:	// CT_Percentage
	case NMSP_DRAWINGML|XML_lumMod:	// CT_Percentage
	case NMSP_DRAWINGML|XML_red:		// CT_Percentage
	case NMSP_DRAWINGML|XML_redOff:	// CT_Percentage
	case NMSP_DRAWINGML|XML_redMod:	// CT_Percentage
	case NMSP_DRAWINGML|XML_green:		// CT_Percentage
	case NMSP_DRAWINGML|XML_greenOff:	// CT_Percentage
	case NMSP_DRAWINGML|XML_greenMod:	// CT_Percentage
	case NMSP_DRAWINGML|XML_blue:		// CT_Percentage
	case NMSP_DRAWINGML|XML_blueOff:	// CT_Percentage
	case NMSP_DRAWINGML|XML_blueMod:	// CT_Percentage
		break;
	}
	return this;
}

} }
