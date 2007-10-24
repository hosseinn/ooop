/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: excelfragmentbase.cxx,v $
 *
 *  $Revision: 1.1.2.10 $
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

#include "oox/xls/excelfragmentbase.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::xml::sax::SAXException;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::com::sun::star::xml::sax::XFastAttributeList;
using ::oox::core::AttributeList;
using ::oox::core::FragmentHandler;
using ::oox::core::Relation;
using ::oox::core::RelationPtr;
using ::oox::core::Relations;
using ::oox::core::XmlFilterRef;

namespace oox {
namespace xls {

// ============================================================================

ExcelFragmentBase::ExcelFragmentBase( const XmlFilterRef& rxFilter, const OUString& rFragmentPath ) :
    FragmentHandler( rxFilter, rFragmentPath )
{
}

// com.sun.star.xml.sax.XFastContextHandler interface -------------------------

Reference< XFastContextHandler > SAL_CALL ExcelFragmentBase::createFastChildContext(
        sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs ) throw( SAXException, RuntimeException )
{
    return implCreateChildContext( nElement, rxAttribs );
}

void SAL_CALL ExcelFragmentBase::startFastElement(
        sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs ) throw( SAXException, RuntimeException )
{
    implStartCurrentContext( nElement, rxAttribs );
}

void SAL_CALL ExcelFragmentBase::characters( const OUString& rChars ) throw( SAXException, RuntimeException )
{
    implCharacters( rChars );
}

void SAL_CALL ExcelFragmentBase::endFastElement( sal_Int32 nElement ) throw( SAXException, RuntimeException )
{
    implEndCurrentContext( nElement );
}

// oox.xls.ContextHelper interface --------------------------------------------

bool ExcelFragmentBase::onCanCreateContext( sal_Int32 )
{
    return false;
}

Reference< XFastContextHandler > ExcelFragmentBase::onCreateContext( sal_Int32, const AttributeList& )
{
    // default behaviour: return this to reuse the same instance
    return this;
}

void ExcelFragmentBase::onStartElement( const AttributeList& )
{
}

void ExcelFragmentBase::onEndElement( const OUString& )
{
}

// other ----------------------------------------------------------------------

RelationPtr ExcelFragmentBase::getRelationByType( const OUString& rType ) const
{
    RelationPtr xRelation;
    if( const Relations* pRelations = getRelations().get() )
        xRelation = pRelations->getRelationByType( rType );
    return xRelation;
}

RelationPtr ExcelFragmentBase::getRelationById( const OUString& rRelId ) const
{
    RelationPtr xRelation;
    if( const Relations* pRelations = getRelations().get() )
        xRelation = pRelations->getRelationById( rRelId );
    return xRelation;
}

OUString ExcelFragmentBase::getFragmentPathByRelation( const Relation& rRelation ) const
{
    return resolveRelativePath( rRelation.msTarget );
}

OUString ExcelFragmentBase::getFragmentPathByType( const OUString& rType ) const
{
    OUString aFragmentPath;
    if( const Relation* pRelation = getRelationByType( rType ).get() )
        aFragmentPath = getFragmentPathByRelation( *pRelation );
    return aFragmentPath;
}

OUString ExcelFragmentBase::getFragmentPathByRelId( const OUString& rRelId ) const
{
    OUString aFragmentPath;
    if( const Relation* pRelation = getRelationById( rRelId ).get() )
        aFragmentPath = getFragmentPathByRelation( *pRelation );
    return aFragmentPath;
}

// ============================================================================

} // namespace xls
} // namespace oox

