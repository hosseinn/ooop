/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: slidetimingcontext.cxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: dr $ $Date: 2007/07/06 09:00:51 $
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

#include "oox/ppt/backgroundproperties.hxx"
#include "oox/ppt/slidefragmenthandler.hxx"
#include "oox/ppt/slidetimingcontext.hxx"
#include "oox/drawingml/shapegroupcontext.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"
#include "oox/ppt/timenodelistcontext.hxx"
#include "buildlistcontext.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::container;

namespace oox { namespace ppt {

SlideTimingContext::SlideTimingContext( const ::oox::core::FragmentHandlerRef& xHandler,  TimeNodePtrList & aTimeNodeList ) throw()
	: Context( xHandler )
		, maTimeNodeList( aTimeNodeList )
{
}

SlideTimingContext::~SlideTimingContext() throw()
{

}

void SlideTimingContext::endFastElement( sal_Int32 /*aElement*/ ) throw ( SAXException, RuntimeException)
{
}


Reference< XFastContextHandler > SlideTimingContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
	case NMSP_PPT|XML_bldLst:
		xRet.set( new BuildListContext( getHandler(), xAttribs, maTimeNodeList ) );
		break;
	case NMSP_PPT|XML_extLst:
		// TODO what to do with extLst
		break;
	case NMSP_PPT|XML_tnLst:
		// timing nodes
	{
		xRet.set( new TimeNodeListContext( getHandler(), maTimeNodeList ) );
	}
	break;

	default:
		break;
	}

	if( !xRet.is() )
		xRet.set(this);

	return xRet;
}

void SAL_CALL SlideTimingContext::endDocument(  ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{

}

} }
