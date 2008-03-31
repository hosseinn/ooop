/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: themefragmenthandler.cxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/23 11:41:49 $
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

#include <com/sun/star/beans/XMultiPropertySet.hpp>
#include <com/sun/star/container/XNamed.hpp>

#include "oox/drawingml/themefragmenthandler.hxx"
#include "oox/drawingml/objectdefaultcontext.hxx"
#include "oox/drawingml/themeelementscontext.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using ::rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::container;

namespace oox { namespace drawingml {

ThemeFragmentHandler::ThemeFragmentHandler( const XmlFilterRef& xFilter, const OUString& rFragmentPath, const ThemePtr pThemePtr )
	throw()
: FragmentHandler( xFilter, rFragmentPath )
, mpThemePtr( pThemePtr )
{
}
ThemeFragmentHandler::~ThemeFragmentHandler()
	throw()
{

}
Reference< XFastContextHandler > ThemeFragmentHandler::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& /*xAttribs */ )
	throw (SAXException, RuntimeException)
{
	// CT_OfficeStyleSheet
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_themeElements:				// CT_BaseStyles
		{
			xRet.set( new themeElementsContext( this, mpThemePtr ) );
			break;
		}
		case NMSP_DRAWINGML|XML_objectDefaults:				// CT_ObjectStyleDefaults
		{
			xRet.set( new objectDefaultContext( this, mpThemePtr ) );
			break;
		}
		case NMSP_DRAWINGML|XML_extraClrSchemeLst:			// CT_ColorSchemeList
		break;
		case NMSP_DRAWINGML|XML_custClrLst:					// CustomColorList
		break;
		case NMSP_DRAWINGML|XML_ext:						// CT_OfficeArtExtension
		break;
	}
	if( !xRet.is() )
		xRet.set( this );
	return xRet;
}
void SAL_CALL ThemeFragmentHandler::endDocument()
	throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
	try
	{
/*
		if( !maSlideProperties.empty() )
		{

			uno::Reference< beans::XMultiPropertySet > xMSet( mxSlide, uno::UNO_QUERY );
			if( xMSet.is() )
			{
				uno::Sequence< OUString > aNames;
				uno::Sequence< uno::Any > aValues;
				maSlideProperties.makeSequence( aNames, aValues );
				xMSet->setPropertyValues( aNames,  aValues);
			}
			else
			{
				uno::Reference< beans::XPropertySet > xSet( mxSlide, uno::UNO_QUERY_THROW );
				for( PropertyMap::const_iterator aIter( maSlideProperties.begin() ); aIter != maSlideProperties.end(); aIter++ )
				{
					xSet->setPropertyValue( (*aIter).first, (*aIter).second );
				}
			}
		}

		Reference< XNamed > xNamed( mxSlide, UNO_QUERY );
		if( xNamed.is() )
			xNamed->setName( maSlideName );
*/
	}
	catch( uno::Exception& )
	{
        OSL_ENSURE( false,
			(rtl::OString("oox::drawingml::ThemeFragmentHandler::EndElement(), "
					"exception caught: ") +
			rtl::OUStringToOString(
				comphelper::anyToString( cppu::getCaughtException() ),
				RTL_TEXTENCODING_UTF8 )).getStr() );
	}
}

//--------------------------------------------------------------------------------------------------------------



} }
