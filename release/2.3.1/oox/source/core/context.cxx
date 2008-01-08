/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: context.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: cl $ $Date: 2007/01/25 16:15:56 $
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

#include "oox/core/context.hxx"

using ::rtl::OUString;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace core {

Context::Context( const FragmentHandlerRef& xHandler )
: mxHandler( xHandler )
{
}

// XFastContextHandler
void Context::startFastElement( ::sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void Context::startUnknownElement( const OUString&, const OUString&, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void Context::endFastElement( ::sal_Int32 ) throw (SAXException, RuntimeException)
{
}

void Context::endUnknownElement( const OUString&, const OUString& ) throw (SAXException, RuntimeException)
{
}

Reference< XFastContextHandler > Context::createFastChildContext( ::sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

Reference< XFastContextHandler > Context::createUnknownChildContext( const OUString&, const OUString&, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

void Context::characters( const OUString& ) throw (SAXException, RuntimeException)
{
}

void Context::ignorableWhitespace( const OUString& ) throw (SAXException, RuntimeException)
{
}

void Context::processingInstruction( const OUString&, const OUString& ) throw (SAXException, RuntimeException)
{
}

} }

