/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: graphicshapecontext.cxx,v $
 *
 *  $Revision: 1.1.2.10 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/21 14:44:02 $
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

#include "oox/drawingml/fillpropertiesgroup.hxx"
#include "oox/drawingml/graphicshapecontext.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/drawingml/drawingmltypes.hxx"
#include "tokens.hxx"
#ifndef _COMPHELPER_PROCESSFACTORY_HXX_
#include <comphelper/processfactory.hxx>
#endif
#include "com/sun/star/container/XNameAccess.hpp"
#include "com/sun/star/io/XStream.hpp"
#include "com/sun/star/beans/XPropertySet.hpp"
#include "com/sun/star/document/XEmbeddedObjectResolver.hpp"

using ::rtl::OUString;
using namespace ::com::sun::star::io;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::beans;
using namespace ::oox::core;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

// ====================================================================
// CT_Picture
GraphicShapeContext::GraphicShapeContext( const FragmentHandlerRef& xHandler, ShapePtr pMasterShapePtr, ShapePtr pShapePtr )
: ShapeContext( xHandler, pMasterShapePtr, pShapePtr )
{
}

Reference< XFastContextHandler > GraphicShapeContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken &(~NMSP_MASK) )
	{
	// CT_ShapeProperties
	case XML_xfrm:
		xRet.set( new Transform2DContext( getHandler(), xAttribs, mpShapePtr ) );
		break;
	case XML_blipFill:
		xRet.set( FillPropertiesGroupContext::StaticCreateContext( getHandler(), (aElementToken&(~NMSP_MASK))|NMSP_DRAWINGML, xAttribs, mpShapePtr->getShapeProperties() ) );
		break;
	}
	if( !xRet.is() )
		xRet.set( ShapeContext::createFastChildContext( aElementToken, xAttribs ) );

	return xRet;
}

// ====================================================================
// CT_GraphicalObjectFrameContext
GraphicalObjectFrameContext::GraphicalObjectFrameContext( const FragmentHandlerRef& xHandler, ShapePtr pMasterShapePtr, ShapePtr pShapePtr )
: ShapeContext( xHandler, pMasterShapePtr, pShapePtr )
{
}

Reference< XFastContextHandler > GraphicalObjectFrameContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken &(~NMSP_MASK) )
	{
	// CT_ShapeProperties
	case XML_nvGraphicFramePr:		// CT_GraphicalObjectFrameNonVisual
		break;
	case XML_xfrm:					// CT_Transform2D
		xRet.set( new Transform2DContext( getHandler(), xAttribs, mpShapePtr ) );
		break;
	case XML_graphic:				// CT_GraphicalObject
		xRet.set( this );
		break;

		case XML_graphicData :			// CT_GraphicalObjectData
		{
			rtl::OUString sUri( xAttribs->getOptionalValue( XML_uri ) );
			if ( sUri == OUString( RTL_CONSTASCII_USTRINGPARAM( "http://schemas.openxmlformats.org/presentationml/2006/ole" ) ) )
				xRet.set( new PresentationOle2006Context( mxHandler, mpShapePtr ) );
			else
				return xRet;
		}
		break;
	}
	if( !xRet.is() )
		xRet.set( ShapeContext::createFastChildContext( aElementToken, xAttribs ) );

	return xRet;
}

// ====================================================================

PresentationOle2006Context::PresentationOle2006Context( const FragmentHandlerRef& xHandler, ShapePtr pShapePtr )
: ShapeContext( xHandler, ShapePtr(), pShapePtr )
{
}

PresentationOle2006Context::~PresentationOle2006Context()
{
	RelationPtr pRelation = getHandler()->getRelations()->getRelationById( msId );
	if( pRelation.get() )
	{
        XmlFilterRef xFilter = getHandler()->getFilter();
		const OUString aFragmentPath( getHandler()->resolveRelativePath( pRelation->msTarget ) );
		Reference< ::com::sun::star::io::XInputStream > xInputStream( xFilter->openInputStream( aFragmentPath ), UNO_QUERY_THROW );

		Sequence< sal_Int8 > aData;
		xInputStream->readBytes( aData, 0x7fffffff );

		::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiServiceFactory > xMSF( xFilter->getModel(), UNO_QUERY );
		Reference< com::sun::star::document::XEmbeddedObjectResolver > xEmbeddedResolver( xMSF->createInstance( OUString::createFromAscii( "com.sun.star.document.ImportEmbeddedObjectResolver" ) ), UNO_QUERY );

		if ( xEmbeddedResolver.is() )
		{
			Reference< com::sun::star::container::XNameAccess > xNA( xEmbeddedResolver, UNO_QUERY );
			if( xNA.is() )
			{
				Reference < XOutputStream > xOLEStream;
				rtl::OUString aURL( RTL_CONSTASCII_USTRINGPARAM( "Obj12345678" ) );
				Any aAny( xNA->getByName( aURL ) );
				aAny >>= xOLEStream;
				if ( xOLEStream.is() )
				{
					xOLEStream->writeBytes( aData );
					xOLEStream->closeOutput();

					const OUString sProtocol(RTL_CONSTASCII_USTRINGPARAM( "vnd.sun.star.EmbeddedObject:" ));
					rtl::OUString aPersistName( xEmbeddedResolver->resolveEmbeddedObjectURL( aURL ) );
					aPersistName = aPersistName.copy( sProtocol.getLength() );

					static const rtl::OUString sPersistName( RTL_CONSTASCII_USTRINGPARAM( "PersistName" ) );
					mpShapePtr->getShapeProperties()[ sPersistName ] <<= aPersistName;
				}
			}
			Reference< XComponent > xComp( xEmbeddedResolver, UNO_QUERY );
			xComp->dispose();
		}
	}
}

Reference< XFastContextHandler > PresentationOle2006Context::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken &(~NMSP_MASK) )
	{
		case XML_oleObj:
		{
			msSpid = xAttribs->getOptionalValue( XML_spid );
			msName = xAttribs->getOptionalValue( XML_name );
			msId = xAttribs->getOptionalValue( NMSP_RELATIONSHIPS|XML_id );
			mnWidth = GetCoordinate( xAttribs->getOptionalValue( XML_imgW ) );
			mnHeight = GetCoordinate( xAttribs->getOptionalValue( XML_imgH ) );
			msProgId = xAttribs->getOptionalValue( XML_progId );
		}
		break;

			case XML_embed:
			{
				mnFollowColorSchemeToken = xAttribs->getOptionalValueToken( XML_followColorScheme, XML_full );
			}
			break;
	}
	if( !xRet.is() )
		xRet.set( ShapeContext::createFastChildContext( aElementToken, xAttribs ) );

	return xRet;
}


} }
