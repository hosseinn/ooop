/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: themeelementscontext.cxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: sj $ $Date: 2007/06/22 15:23:24 $
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

#include "oox/drawingml/themeelementscontext.hxx"
#include "oox/drawingml/clrschemecontext.hxx"
#include "oox/drawingml/fillpropertiesgroup.hxx"
#include "oox/drawingml/lineproperties.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

class fillStyleListContext : public ::oox::core::Context
{
	std::vector< oox::core::PropertyMap	>& mrFillStyleList;
public:
	fillStyleListContext( const ::oox::core::FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rFillStyleList );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
};
fillStyleListContext::fillStyleListContext( const FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rFillStyleList )
: oox::core::Context( xHandler )
, mrFillStyleList( rFillStyleList )
{
}
Reference< XFastContextHandler > fillStyleListContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_noFill:
		case NMSP_DRAWINGML|XML_solidFill:
		case NMSP_DRAWINGML|XML_gradFill:
		case NMSP_DRAWINGML|XML_blipFill:
		case NMSP_DRAWINGML|XML_pattFill:
		case NMSP_DRAWINGML|XML_grpFill:
		{
			mrFillStyleList.push_back( oox::core::PropertyMap() );
			FillPropertiesGroupContext::StaticCreateContext( getHandler(), aElementToken, xAttribs, mrFillStyleList.back() );
		}
		break;
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// ---------------------------------------------------------------------

class lineStyleListContext : public ::oox::core::Context
{
	std::vector< oox::core::PropertyMap	>& mrLineStyleList;
public:
	lineStyleListContext( const ::oox::core::FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rLineStyleList );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
};
lineStyleListContext::lineStyleListContext( const FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rLineStyleList )
: oox::core::Context( xHandler )
, mrLineStyleList( rLineStyleList )
{
}
Reference< XFastContextHandler > lineStyleListContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_ln:
		{
			mrLineStyleList.push_back( oox::core::PropertyMap() );
			xRet.set( new LinePropertiesContext( getHandler(), xAttribs, mrLineStyleList.back() ) );
		}
		break;
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// ---------------------------------------------------------------------

class effectStyleListContext : public ::oox::core::Context
{
	std::vector< oox::core::PropertyMap	>& mrEffectStyleList;
public:
	effectStyleListContext( const ::oox::core::FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap >& rEffectStyleList );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
};
effectStyleListContext::effectStyleListContext( const FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rEffectStyleList )
: oox::core::Context( xHandler )
, mrEffectStyleList( rEffectStyleList )
{
}
Reference< XFastContextHandler > effectStyleListContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& /* xAttribs */ ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_effectStyle:
		{
			mrEffectStyleList.push_back( oox::core::PropertyMap() );
			// todo: last effect list entry needs to be filled/
			break;
		}
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// ---------------------------------------------------------------------

class bgFillStyleListContext : public ::oox::core::Context
{
	std::vector< oox::core::PropertyMap	>& mrBgFillStyleList;
public:
	bgFillStyleListContext( const ::oox::core::FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap >& rBgFillStyleList );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
};
bgFillStyleListContext::bgFillStyleListContext( const FragmentHandlerRef& xHandler, std::vector< oox::core::PropertyMap	>& rBgFillStyleList )
: oox::core::Context( xHandler )
, mrBgFillStyleList( rBgFillStyleList )
{
}
Reference< XFastContextHandler > bgFillStyleListContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_noFill:
		case NMSP_DRAWINGML|XML_solidFill:
		case NMSP_DRAWINGML|XML_gradFill:
		case NMSP_DRAWINGML|XML_blipFill:
		case NMSP_DRAWINGML|XML_pattFill:
		case NMSP_DRAWINGML|XML_grpFill:
		{
			mrBgFillStyleList.push_back( oox::core::PropertyMap() );
			FillPropertiesGroupContext::StaticCreateContext( getHandler(), aElementToken, xAttribs, mrBgFillStyleList.back() );
		}
		break;
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// ---------------------------------------------------------------------

themeElementsContext::themeElementsContext( const ::oox::core::FragmentHandlerRef& xHandler, const ::oox::drawingml::ThemePtr pThemePtr )
: Context( xHandler )
, mpThemePtr( pThemePtr )
{
}

Reference< XFastContextHandler > themeElementsContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	// CT_BaseStyles
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_clrScheme:	// CT_ColorScheme
		{
			xRet.set( new clrSchemeContext( getHandler(), mpThemePtr->getClrScheme() ) );
			break;
		}

		case NMSP_DRAWINGML|XML_fontScheme:	// CT_FontScheme
			break;

		case NMSP_DRAWINGML|XML_fmtScheme:	// CT_StyleMatrix
			mpThemePtr->getStyleName() = xAttribs->getOptionalValue( XML_name );
		break;

			case NMSP_DRAWINGML|XML_fillStyleLst:	// CT_FillStyleList
				xRet.set( new fillStyleListContext( getHandler(), mpThemePtr->getFillStyleList() ) );
			break;
			case NMSP_DRAWINGML|XML_lineStyleLst:	// CT_LineStyleList
				xRet.set( new lineStyleListContext( getHandler(), mpThemePtr->getLineStyleList() ) );
			break;
			case NMSP_DRAWINGML|XML_effectStyleLst:	// CT_EffectStyleList
				xRet.set( new effectStyleListContext( getHandler(), mpThemePtr->getEffectStyleList() ) );
			break;
			case NMSP_DRAWINGML|XML_bgFillStyleLst:	// CT_BackgroundFillStyleList
				xRet.set( new bgFillStyleListContext( getHandler(), mpThemePtr->getBgFillStyleList() ) );
			break;

		case NMSP_DRAWINGML|XML_extLst:		// CT_OfficeArtExtensionList
		break;
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}

} }
