/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: fillpropertiesgroup.cxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: sj $ $Date: 2007/06/13 12:50:36 $
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

#include "comphelper/anytostring.hxx"
#include "cppuhelper/exc_hlp.hxx"
#include "oox/drawingml/colorchoicecontext.hxx"
#include "oox/drawingml/fillpropertiesgroup.hxx"
#include <com/sun/star/graphic/XGraphicProvider.hpp>
#include <com/sun/star/drawing/BitmapMode.hpp>
#include "oox/core/propertymap.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using ::rtl::OUString;
using ::com::sun::star::beans::NamedValue;
using namespace ::oox::core;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::io;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::graphic;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

// ---------------------------------------------------------------------

class NoFillContext : public FillPropertiesGroupContext
{
public:
	NoFillContext( const oox::core::FragmentHandlerRef& xHandler, ::oox::core::PropertyMap& rProperties ) throw();
};

// ---------------------------------------------------------------------

class SolidColorFillPropertiesContext : public FillPropertiesGroupContext
{
public:
	SolidColorFillPropertiesContext( const oox::core::FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, ::oox::core::PropertyMap& rProperties ) throw();

	virtual void SAL_CALL startFastElement( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
	virtual void SAL_CALL endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException);
	virtual Reference< XFastContextHandler > SAL_CALL createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);

private:
	sal_Int32 mnColor;
};

// ---------------------------------------------------------------------

class GradFillPropertiesContext : public FillPropertiesGroupContext
{
public:
	GradFillPropertiesContext( const oox::core::FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, ::oox::core::PropertyMap& rProperties ) throw();

	virtual void SAL_CALL endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException);
	virtual Reference< XFastContextHandler > SAL_CALL createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
};

// ---------------------------------------------------------------------

class BlipFillPropertiesContext : public FillPropertiesGroupContext
{
public:
	BlipFillPropertiesContext( const oox::core::FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, ::oox::core::PropertyMap& rProperties ) throw();

	virtual void SAL_CALL endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException);
	virtual Reference< XFastContextHandler > SAL_CALL createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);

private:
	BitmapMode	meBitmapMode;
	sal_Int32	mnWidth, mnHeight;
	OUString	msEmbed;
	OUString	msLink;
};

// ---------------------------------------------------------------------

class PattFillPropertiesContext : public FillPropertiesGroupContext
{
public:
	PattFillPropertiesContext( const oox::core::FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, ::oox::core::PropertyMap& rProperties ) throw();

	virtual void SAL_CALL startFastElement( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
	virtual void SAL_CALL endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException);
	virtual Reference< XFastContextHandler > SAL_CALL createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
};

// ---------------------------------------------------------------------

class GrpFillPropertiesContext : public FillPropertiesGroupContext
{
public:
	GrpFillPropertiesContext( const oox::core::FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, ::oox::core::PropertyMap& rProperties ) throw();

	virtual void SAL_CALL startFastElement( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
	virtual void SAL_CALL endFastElement( sal_Int32 aElementToken ) throw (SAXException, RuntimeException);
	virtual Reference< XFastContextHandler > SAL_CALL createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException);
};

// ---------------------------------------------------------------------

FillPropertiesGroupContext::FillPropertiesGroupContext( const FragmentHandlerRef& xHandler, ::com::sun::star::drawing::FillStyle eFillStyle, PropertyMap& rProperties ) throw()
: ::oox::core::Context( xHandler )
, mrProperties( rProperties )
{
	static const OUString sFillStyle( RTL_CONSTASCII_USTRINGPARAM( "FillStyle" ) );
	mrProperties[ sFillStyle ] <<= eFillStyle;
}

// ---------------------------------------------------------------------

Reference< XFastContextHandler > FillPropertiesGroupContext::StaticCreateContext( const FragmentHandlerRef& xHandler, sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs, ::oox::core::PropertyMap& rProperties ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
	case NMSP_DRAWINGML|XML_noFill:
		xRet.set( new NoFillContext( xHandler, rProperties ) );
		break;
	case NMSP_DRAWINGML|XML_solidFill:
		xRet.set( new SolidColorFillPropertiesContext( xHandler, xAttribs, rProperties ) );
		break;
	case NMSP_DRAWINGML|XML_gradFill:
		xRet.set( new GradFillPropertiesContext( xHandler, xAttribs, rProperties ) );
		break;
	case NMSP_DRAWINGML|XML_blipFill:
		xRet.set( new BlipFillPropertiesContext( xHandler, xAttribs, rProperties ) );
		break;
	case NMSP_DRAWINGML|XML_pattFill:
		xRet.set( new PattFillPropertiesContext( xHandler, xAttribs, rProperties ) );
		break;
	case NMSP_DRAWINGML|XML_grpFill:
		xRet.set( new GrpFillPropertiesContext( xHandler, xAttribs, rProperties ) );
		break;
	}
	return xRet;
}

// ---------------------------------------------------------------------

NoFillContext::NoFillContext( const FragmentHandlerRef& xHandler, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_NONE, rProperties )
{
}

// ---------------------------------------------------------------------

SolidColorFillPropertiesContext::SolidColorFillPropertiesContext( const FragmentHandlerRef& xHandler, const Reference< XFastAttributeList >&, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_SOLID, rProperties )
, mnColor( 0 )
{
}

void SolidColorFillPropertiesContext::startFastElement( sal_Int32 /* aElementToken */, const Reference< XFastAttributeList >& /* xAttribs */ ) throw (SAXException, RuntimeException)
{
}

void SolidColorFillPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
	static const OUString sFillColor( RTL_CONSTASCII_USTRINGPARAM( "FillColor" ) );
	mrProperties[ sFillColor ]<<= mnColor;
}

Reference< XFastContextHandler > SolidColorFillPropertiesContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	// colorTransformGroup

	// color should be available as rgb in member mnColor already, now modify it depending on
	// the transformation elements

	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{

	case NMSP_DRAWINGML|XML_scrgbClr:	// CT_ScRgbColor
	case NMSP_DRAWINGML|XML_srgbClr:	// CT_SRgbColor
	case NMSP_DRAWINGML|XML_hslClr:	// CT_HslColor
	case NMSP_DRAWINGML|XML_sysClr:	// CT_SystemColor
	case NMSP_DRAWINGML|XML_schemeClr:	// CT_SchemeColor
	case NMSP_DRAWINGML|XML_prstClr:	// CT_PresetColor
		{
            xRet.set( new colorChoiceContext( getHandler(), mnColor ) );
			break;
		}
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
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// ---------------------------------------------------------------------

GradFillPropertiesContext::GradFillPropertiesContext( const FragmentHandlerRef& xHandler, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >&, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_GRADIENT, rProperties )
{
}

void GradFillPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
}

Reference< XFastContextHandler > GradFillPropertiesContext::createFastChildContext( sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	return this;
}

// ---------------------------------------------------------------------
// CT_BlipFill
// ---------------------------------------------------------------------

BlipFillPropertiesContext::BlipFillPropertiesContext( const FragmentHandlerRef& xHandler, const Reference< XFastAttributeList >&, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_BITMAP, rProperties )
, meBitmapMode( BitmapMode_REPEAT )
, mnWidth( 0 )
, mnHeight( 0 )
{
	/* todo
	if( xAttribs->hasAttribute( XML_dpi ) )
	{
		 xsd:unsignedInt
		DPI (dots per inch) used to calculate the size of the blip. If not present or zero,
					the DPI in the blip is used.

	}
	if( xAttribs->hasAttribute( XML_rotWithShape ) )
	{
		xsd:boolean
		fill should rotate with the shape
	}
	*/
}

void BlipFillPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
	if( msEmbed.getLength() )
	{
		RelationPtr pRelation = getHandler()->getRelations()->getRelationById( msEmbed );
		if( pRelation.get() ) try
		{
			// get the input stream for the fill bitmap
            XmlFilterRef xFilter = getHandler()->getFilter();
			const OUString aFragmentPath( getHandler()->resolveRelativePath( pRelation->msTarget ) );
            Reference< XInputStream > xInputStream( xFilter->openInputStream( aFragmentPath ), UNO_QUERY_THROW );

			// load the fill bitmap into an XGraphic with the GraphicProvider
			static const OUString sGraphicProvider( RTL_CONSTASCII_USTRINGPARAM( "com.sun.star.graphic.GraphicProvider" ) );
            Reference< XMultiServiceFactory > xMSFT( xFilter->getServiceFactory(), UNO_QUERY_THROW );
			Reference< XGraphicProvider > xGraphicProvider( xMSFT->createInstance( sGraphicProvider ), UNO_QUERY_THROW );

			static const OUString sInputStream( RTL_CONSTASCII_USTRINGPARAM( "InputStream" ) );
			PropertyValues aMediaProperties(1);
			aMediaProperties[0].Name = sInputStream;
			aMediaProperties[0].Value <<= xInputStream;

			Reference< XGraphic > xGraphic( xGraphicProvider->queryGraphic(aMediaProperties ) );

			// insert the FillBitmap property with the imported graphic
			static const OUString sFillBitmap( RTL_CONSTASCII_USTRINGPARAM( "FillBitmap" ) );
			mrProperties[ sFillBitmap ] <<= xGraphic;
		}
		catch( Exception& )
		{
            OSL_ENSURE( false,
				(rtl::OString("oox::drawingml::BlipFillPropertiesContext::EndElement(), "
						"exception caught: ") +
				rtl::OUStringToOString(
					comphelper::anyToString( cppu::getCaughtException() ),
					RTL_TEXTENCODING_UTF8 )).getStr() );

		}

		static const OUString sFillBitmapMode( RTL_CONSTASCII_USTRINGPARAM( "FillBitmapMode" ) );
		mrProperties[ sFillBitmapMode ] <<= meBitmapMode;
	}
}

Reference< XFastContextHandler > BlipFillPropertiesContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	switch( aElementToken )
	{
	case NMSP_DRAWINGML|XML_blip:			// CT_Blip
//		mnWidth = xAttribs.getInt32( XML_w );
//		mnHeight = xAttribs.getInt32( XML_h );
		msEmbed = xAttribs->getOptionalValue(NMSP_RELATIONSHIPS|XML_embed);	// relationship identifer for embedded blobs
		msLink = xAttribs->getOptionalValue(NMSP_RELATIONSHIPS|XML_link);	// relationship identifer for linked blobs
		break;
	case NMSP_DRAWINGML|XML_srcRect:		// CT_RelativeRect
// todo		maSrcRect = GetRelativeRect( xAttribs );
		break;
	case NMSP_DRAWINGML|XML_tile:			// CT_TileInfo
		meBitmapMode = BitmapMode_REPEAT;
/* todo
		mnTileX = GetCoordinate( xAttribs[ XML_tx ] );	// additional horizontal offset after alignment
		mnTileY = GetCoordinate( xAttribs[ XML_ty ] );	// additional vertical offset after alignment
		mnSX = xAttribs.getInt32( XML_sx, 100000 );		// amount to horizontally scale the srcRect
		mnSX = xAttribs.getInt32( XML_sy, 100000 );		// amount to vertically scale the srcRect
		mnFlip = xAttribs->getOptionalValueToken( XML_flip ); // ST_TileFlipMode how to flip the tile when it repeats
		mnAlign = xAttribs->getOptionalValueToken( XML_algn ); // ST_RectAlignment where to align the first tile with respect to the shape; alignment happens after the scaling, but before the additional offset.
*/
		break;
	case NMSP_DRAWINGML|XML_stretch:		// CT_StretchInfo
		meBitmapMode = BitmapMode_STRETCH;
		break;
	case NMSP_DRAWINGML|XML_fillRect:
// todo		maFillRect = GetRelativeRect( xAttribs );
		break;
	}
	return this;
}

// ---------------------------------------------------------------------

PattFillPropertiesContext::PattFillPropertiesContext( const FragmentHandlerRef& xHandler, const Reference< XFastAttributeList >&, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_HATCH, rProperties )
{
}

void PattFillPropertiesContext::startFastElement( sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void PattFillPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
}

Reference< XFastContextHandler > PattFillPropertiesContext::createFastChildContext( sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	return this;
}

// ---------------------------------------------------------------------

GrpFillPropertiesContext::GrpFillPropertiesContext( const FragmentHandlerRef& xHandler,const Reference< XFastAttributeList >&, PropertyMap& rProperties ) throw()
: FillPropertiesGroupContext( xHandler, FillStyle_NONE, rProperties )
{
}

void GrpFillPropertiesContext::startFastElement( sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void GrpFillPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
}

Reference< XFastContextHandler > GrpFillPropertiesContext::createFastChildContext( sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	return this;
}

} }
