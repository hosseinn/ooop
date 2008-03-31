/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: layoutfragmenthandler.cxx,v $
 *
 *  $Revision: 1.1.2.10 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/30 12:15:26 $
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

#include "oox/ppt/layoutfragmenthandler.hxx"
#include "oox/drawingml/shapegroupcontext.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::container;

namespace oox { namespace ppt {

// CT_SlideLayout

LayoutFragmentHandler::LayoutFragmentHandler( const oox::core::XmlFilterRef& xFilter, const ::rtl::OUString& rFragmentPath, oox::ppt::SlidePersistPtr pMasterPersistPtr )
	throw()
: SlideFragmentHandler( xFilter, rFragmentPath, pMasterPersistPtr, Layout )
{
}

LayoutFragmentHandler::~LayoutFragmentHandler()
	throw()
{

}

Reference< XFastContextHandler > LayoutFragmentHandler::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs )
	throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet( this );
	switch( aElementToken )
	{
		case NMSP_PPT|XML_sldLayout:		// CT_SlideLayout
			mpSlidePersistPtr->setLayoutValueToken( xAttribs->getOptionalValueToken( XML_type, 0 ) );	// CT_SlideLayoutType
		break;
		default:
			xRet.set( SlideFragmentHandler::createFastChildContext( aElementToken, xAttribs ) );
	}
	return xRet;
}

void SAL_CALL LayoutFragmentHandler::endDocument()
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
			(rtl::OString("oox::ppt::SlideFragmentHandler::EndElement(), "
					"exception caught: ") +
			rtl::OUStringToOString(
				comphelper::anyToString( cppu::getCaughtException() ),
				RTL_TEXTENCODING_UTF8 )).getStr() );
	}
}

} }
