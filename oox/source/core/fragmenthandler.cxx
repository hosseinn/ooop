/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: fragmenthandler.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/23 14:08:44 $
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

#include "oox/core/fragmenthandler.hxx"

using ::rtl::OUString;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace core {

FragmentHandler::FragmentHandler( const oox::core::XmlFilterRef& xFilter, const OUString& rFragmentPath )
: mxFilter( xFilter )
{
    mxRelations = mxFilter->getRelations( rFragmentPath );
}

// --------------------------------------------------------------------

rtl::OUString FragmentHandler::resolveRelativePath( const OUString& rRelPath ) const
{
    OSL_ENSURE( mxLocator.is() , "oox::core::FragmentHandler::resolveRelativePath(), no locator!" );
	OUString sSystemId;
	if( mxLocator.is() )
		sSystemId = mxLocator->getSystemId();
	return resolveRelativePath( sSystemId, rRelPath );
}

// --------------------------------------------------------------------

OUString FragmentHandler::resolveRelativePath( const OUString& rFragmentPath, const OUString& rRelPath )
{
	static const sal_Unicode nSlash( '/' );

	if( (rRelPath.getLength() && rRelPath[0] == nSlash) || (rFragmentPath.getLength() == 0) )
		return rRelPath;

	OUString aPath( rFragmentPath );
	sal_Int32 n = aPath.lastIndexOf(nSlash);
	if( n != -1 )
		aPath = aPath.copy( 0, n );

	const OUString sBack( CREATE_OUSTRING("../") );

    // First, count the number of "../"'s found in relative path string.
    sal_Int32 nCount = 0, nPos = 0;
    while ( true )
    {
        nPos = rRelPath.indexOf(sBack, nCount*3);
        if ( nPos != nCount*3 )
            break;
        ++nCount;
    }

    // Now, reduce the base path's directory level by the count.
    for ( sal_Int32 i = 0; i < nCount; ++i )
    {
        sal_Int32 pos = aPath.lastIndexOf(nSlash);
        if ( pos == -1 )
            // This is unexpected.  Bail out.
            return rRelPath;
        aPath = aPath.copy(0, pos);
    }

	aPath += OUString( nSlash );
	aPath += rRelPath.copy(nCount*3);
	return aPath;
}

void FragmentHandler::startDocument(  ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::endDocument(  ) throw (SAXException, RuntimeException)
{
}

Reference< XFastContextHandler > FragmentHandler::createFastChildContext( ::sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

Reference< XFastContextHandler > FragmentHandler::createUnknownChildContext( const OUString&, const OUString&, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

void FragmentHandler::setDocumentLocator( const Reference< XLocator >& xLocator ) throw (SAXException, RuntimeException)
{
	mxLocator = xLocator;
}

// XFastContextHandler
void FragmentHandler::startFastElement( ::sal_Int32, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::startUnknownElement( const OUString&, const OUString&, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::endFastElement( ::sal_Int32 ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::endUnknownElement( const OUString&, const OUString& ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::characters( const OUString& ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::ignorableWhitespace( const OUString& ) throw (SAXException, RuntimeException)
{
}

void FragmentHandler::processingInstruction( const OUString&, const OUString& ) throw (SAXException, RuntimeException)
{
}

} }
