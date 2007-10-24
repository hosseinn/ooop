/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: lineproperties.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: sj $ $Date: 2007/06/25 16:16:41 $
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

#include <com/sun/star/drawing/LineJoint.hpp>
#include <com/sun/star/drawing/LineStyle.hpp>
#include "oox/drawingml/colorchoicecontext.hxx"
#include "oox/drawingml/lineproperties.hxx"
#include "oox/drawingml/drawingmltypes.hxx"
#include "oox/core/propertymap.hxx"
#include "oox/core/namespaces.hxx"

#include "tokens.hxx"

using ::rtl::OUString;
using ::com::sun::star::beans::NamedValue;
using namespace ::oox::core;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

// CT_LineProperties

namespace oox { namespace drawingml {
// ---------------------------------------------------------------------

LinePropertiesContext::LinePropertiesContext( const FragmentHandlerRef& xHandler, const Reference< XFastAttributeList >& xAttribs, PropertyMap& rProperties ) throw()
: Context( xHandler )
, mrProperties( rProperties )
, mnLineColor( 0 )
, mbLineColorUsed( sal_False )
{
	// ST_LineWidth
	if( xAttribs->hasAttribute( XML_w ) )
	{
		static const OUString sLineWidth( RTL_CONSTASCII_USTRINGPARAM( "LineWidth" ) );
		sal_Int32 nWidth = GetCoordinate( xAttribs->getOptionalValue( XML_w ) );
		mrProperties[ sLineWidth ] <<= nWidth;
	}

	// "ST_LineCap"
	/* todo: unsuported
	if( xAttribs->hasAttribute( XML_cap ) )
	{
		switch( xAttribs->getOptionalValueToken( XML_CAP, FastToken::DONTKNOW ) )
		{
		case XML_rnd:	// Rounded ends. Semi-circle protrudes by half line width.
		case XML_sq:	// Square protrudes by half line width.
		case XML_flat:	// Line ends at end point.
		default:
			DBG_ERROR("oox::drawingml::LinePropertiesContext::LinePropertiesContext(), unknown line cap style");
			break;
		}
	}
	*/
	// if ( xAttribs->hasAttribute( XML_cmpd ) )	ST_CompoundLine
	// if ( xAttribs->hasAttribute( XML_algn ) )	ST_PenAlignment

}

LinePropertiesContext::~LinePropertiesContext()
{
	if ( mbLineColorUsed )
	{
		const rtl::OUString sLineColor( RTL_CONSTASCII_USTRINGPARAM( "LineColor" ) );
		mrProperties[ sLineColor ] <<= mnLineColor;
	}
}

Reference< XFastContextHandler > LinePropertiesContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	const rtl::OUString sLineStyle( RTL_CONSTASCII_USTRINGPARAM( "LineStyle" ) );
	switch( aElementToken )
	{
		// LineFillPropertiesGroup, line fillings currently unsuported
		case NMSP_DRAWINGML|XML_noFill:
			mrProperties[ sLineStyle ] <<= LineStyle_NONE;
		break;
		case NMSP_DRAWINGML|XML_solidFill:
		{
			mbLineColorUsed = sal_True;
			mrProperties[ sLineStyle ] <<= LineStyle_SOLID;
			xRet = new colorChoiceContext( getHandler(), mnLineColor );
		}
		break;
		case NMSP_DRAWINGML|XML_gradFill:
		case NMSP_DRAWINGML|XML_pattFill:
			mrProperties[ sLineStyle ] <<= LineStyle_SOLID;
		break;

		// LineDashPropertiesGroup
		case NMSP_DRAWINGML|XML_prstDash:	// CT_PresetLineDashProperties
		case NMSP_DRAWINGML|XML_custDash:	// CT_DashStopList
		break;

		// LineJoinPropertiesGroup
		case NMSP_DRAWINGML|XML_round:
		case NMSP_DRAWINGML|XML_bevel:
		case NMSP_DRAWINGML|XML_miter:
		{
			LineJoint eJoint = (aElementToken == (NMSP_DRAWINGML|XML_round)) ? LineJoint_ROUND : 
								(aElementToken == (NMSP_DRAWINGML|XML_bevel)) ? LineJoint_BEVEL :
									LineJoint_MITER;
			static const OUString sLineJoint( RTL_CONSTASCII_USTRINGPARAM( "LineJoint" ) );
			mrProperties[ sLineJoint ] <<= eJoint;
		}
		break;

		case NMSP_DRAWINGML|XML_headEnd:	// CT_LineEndProperties
		case NMSP_DRAWINGML|XML_tailEnd:	// CT_LineEndProperties
		{
			if( xAttribs->hasAttribute( XML_type ) )	// ST_LineEndType
			{
				sal_Int32 nType = xAttribs->getOptionalValueToken( XML_type, 0 );
				switch( nType )
				{
					case XML_none:
					case XML_triangle:
					case XML_stealth:
					case XML_diamond:
					case XML_oval:
					case XML_arrow:
						break;
				}

				/* todo either create a polypolygon or use one of our default line endings by name
				static const OUString sLineStartName( RTL_CONSTASCII_USTRINGPARAM( "LineStartName" ) );
				static const OUString sLineEndName( RTL_CONSTASCII_USTRINGPARAM( "LineEndName" ) );
				mrProperties.push_back(NamedValue( );
				*/
			}
		}
		break;
	}
	if ( !xRet.is() )
		xRet.set( this );
	return xRet;
}

} }
