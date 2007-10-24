/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: relationshandler.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/02/27 12:49:41 $
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

#include <com/sun/star/xml/sax/FastToken.hpp>

#include "oox/core/relationshandler.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/context.hxx"

#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace core {

// --------------------------------------------------------------------

RelationsFragmentHandler::RelationsFragmentHandler( const rtl::OUString& /*rFragmentPath*/, RelationsRef& rxRelations )
: mrxRelations( rxRelations )
{
}

// --------------------------------------------------------------------
// XFastDocumentHandler

Reference< XFastContextHandler > RelationsFragmentHandler::createFastChildContext( ::sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttributes ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
	case NMSP_PACKAGE_RELATIONSHIPS|XML_Relationship:
	{
		OUString sId( xAttributes->getOptionalValue(XML_Id) );
		OUString sTarget( xAttributes->getOptionalValue(XML_Target) );
		OUString sType( xAttributes->getOptionalValue(XML_Type) );
        OSL_ENSURE( sId.getLength() && sType.getLength() && sTarget.getLength(), "oox::core::RelationsContext::CreateChildContext(), incomplete relation element?" );
		if( sId.getLength() && sType.getLength() && sTarget.getLength() )
			mrxRelations->push_back( RelationPtr( new Relation( sId, sType, sTarget ) ) );
		break;
	}
	case NMSP_PACKAGE_RELATIONSHIPS|XML_Relationships:
	{
		xRet.set( this );
	}
	}

	return xRet;
}

// --------------------------------------------------------------------

void RelationsFragmentHandler::startDocument(  ) throw (SAXException, RuntimeException)
{
}

// --------------------------------------------------------------------

void RelationsFragmentHandler::endDocument(  ) throw (SAXException, RuntimeException)
{
}

// --------------------------------------------------------------------

Reference< XFastContextHandler > RelationsFragmentHandler::createUnknownChildContext( const ::rtl::OUString&, const ::rtl::OUString&, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

// --------------------------------------------------------------------

void RelationsFragmentHandler::setDocumentLocator( const Reference< XLocator >& ) throw (SAXException, RuntimeException)
{
}

// XFastContextHandler
void SAL_CALL RelationsFragmentHandler::startFastElement( ::sal_Int32, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::startUnknownElement( const ::rtl::OUString&, const ::rtl::OUString&, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::endFastElement( ::sal_Int32 ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::endUnknownElement( const ::rtl::OUString&, const ::rtl::OUString& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::characters( const ::rtl::OUString& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::ignorableWhitespace( const ::rtl::OUString& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

void SAL_CALL RelationsFragmentHandler::processingInstruction( const ::rtl::OUString&, const ::rtl::OUString& ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
}

} }
