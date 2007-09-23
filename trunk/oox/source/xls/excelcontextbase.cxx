/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: excelcontextbase.cxx,v $
 *
 *  $Revision: 1.1.2.12 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:48 $
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

#include "oox/xls/excelcontextbase.hxx"
#include "oox/xls/excelfragmentbase.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::xml::sax::SAXException;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::com::sun::star::xml::sax::XFastAttributeList;
using ::oox::core::AttributeList;
using ::oox::core::Context;
using ::oox::core::FragmentHandlerRef;

namespace oox {
namespace xls {

// ============================================================================

ExcelContextBase::ExcelContextBase( const ExcelFragmentBase& rParent ) :
    Context( const_cast< ExcelFragmentBase* >( &rParent ) )
{
}

ExcelContextBase::ExcelContextBase( const ExcelContextBase& rParent ) :
    Context( rParent ),
    ContextHelper()
{
}

// com.sun.star.xml.sax.XFastContextHandler interface -------------------------

Reference< XFastContextHandler > SAL_CALL ExcelContextBase::createFastChildContext(
        sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs ) throw( SAXException, RuntimeException )
{
    return implCreateChildContext( nElement, rxAttribs );
}

void SAL_CALL ExcelContextBase::startFastElement(
        sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs ) throw( SAXException, RuntimeException )
{
    implStartCurrentContext( nElement, rxAttribs );
}

void SAL_CALL ExcelContextBase::characters( const OUString& rChars ) throw( SAXException, RuntimeException )
{
    implCharacters( rChars );
}

void SAL_CALL ExcelContextBase::endFastElement( sal_Int32 nElement ) throw( SAXException, RuntimeException )
{
    implEndCurrentContext( nElement );
}

// oox.xls.ContextHelper interface --------------------------------------------

bool ExcelContextBase::onCanCreateContext( sal_Int32 )
{
    return false;
}

Reference< XFastContextHandler > ExcelContextBase::onCreateContext( sal_Int32, const AttributeList& )
{
    // default behaviour: return this to reuse the same instance
    return this;
}

void ExcelContextBase::onStartElement( const AttributeList& )
{
}

void ExcelContextBase::onEndElement( const OUString& )
{
}

// ============================================================================

} // namespace xls
} // namespace oox

