/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: shapepropertiescontext.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: sj $ $Date: 2007/06/29 14:47:17 $
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

#ifndef _TOOLS_DEBUG_HXX
#include <tools/debug.hxx>
#endif

#include <com/sun/star/xml/sax/FastToken.hpp>
#include <com/sun/star/drawing/LineStyle.hpp>
#include <com/sun/star/beans/XMultiPropertySet.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/container/XNamed.hpp>

#include "oox/drawingml/shapepropertiescontext.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/drawingml/fillpropertiesgroup.hxx"
#include "oox/drawingml/lineproperties.hxx"
#include "oox/drawingml/drawingmltypes.hxx"
#include "oox/drawingml/customshapegeometry.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace oox::core;
using namespace ::com::sun::star;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

// ====================================================================

// CT_ShapeProperties
ShapePropertiesContext::ShapePropertiesContext( const ContextRef& xParent, ::oox::drawingml::ShapePtr pShapePtr )
: Context( xParent->getHandler() )
, mxParent( xParent )
, mpShapePtr( pShapePtr )
{
}

// --------------------------------------------------------------------

void ShapePropertiesContext::endFastElement( sal_Int32 aElementToken ) throw( SAXException, RuntimeException )
{
	mxParent->endFastElement( aElementToken );
}

// --------------------------------------------------------------------

Reference< XFastContextHandler > ShapePropertiesContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
	// CT_Transform2D
	case NMSP_DRAWINGML|XML_xfrm:
		xRet.set( new Transform2DContext( getHandler(), xAttribs, mpShapePtr ) );
		break;

	// GeometryGroup
	case NMSP_DRAWINGML|XML_custGeom:	// custom geometry "CT_CustomGeometry2D"
	case NMSP_DRAWINGML|XML_prstGeom:	// preset geometry "CT_PresetGeometry2D"
		xRet.set( new CustomShapeGeometryContext( getHandler(), aElementToken, xAttribs, mpShapePtr->getCustomShapeGeometry() ) );
		break;

	// CT_LineProperties
	case NMSP_DRAWINGML|XML_ln:
		xRet.set( new LinePropertiesContext( getHandler(), xAttribs, mpShapePtr->getShapeProperties() ) );
		break;

	// EffectPropertiesGroup
	// todo not supported by core
	case NMSP_DRAWINGML|XML_effectLst:	// CT_EffectList
	case NMSP_DRAWINGML|XML_effectDag:	// CT_EffectContainer
		break;
		
	// todo
	case NMSP_DRAWINGML|XML_scene3d:	// CT_Scene3D
	case NMSP_DRAWINGML|XML_sp3d:		// CT_Shape3D
		break;
	}

	// FillPropertiesGroup
	if( !xRet.is() )
		xRet.set( FillPropertiesGroupContext::StaticCreateContext( getHandler(), aElementToken, xAttribs, mpShapePtr->getShapeProperties() ) );

	return xRet;
}

} }