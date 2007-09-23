/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: excelfragmentbase.hxx,v $
 *
 *  $Revision: 1.1.2.11 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:58:00 $
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

#ifndef OOX_XLS_EXCELFRAGMENTBASE_HXX
#define OOX_XLS_EXCELFRAGMENTBASE_HXX

#include "oox/core/fragmenthandler.hxx"
#include "oox/xls/contexthelper.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/worksheethelper.hxx"

namespace oox {
namespace xls {

#define CREATE_RELATIONS_TYPE( ascii ) CREATE_OUSTRING( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" ascii )

// ============================================================================

class ExcelFragmentBase : public ::oox::core::FragmentHandler, public ContextHelper
{
public:
    explicit            ExcelFragmentBase(
                            const ::oox::core::XmlFilterRef& rxFilter,
                            const ::rtl::OUString& rFragmentPath );

    // com.sun.star.xml.sax.XFastContextHandler interface ---------------------

    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL
                        createFastChildContext(
                            sal_Int32 nElement,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& rxAttribs )
                            throw(  ::com::sun::star::xml::sax::SAXException,
                                    ::com::sun::star::uno::RuntimeException );

    virtual void SAL_CALL startFastElement(
                            sal_Int32 nElement,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& rxAttribs )
                            throw(  ::com::sun::star::xml::sax::SAXException,
                                    ::com::sun::star::uno::RuntimeException );

    virtual void SAL_CALL characters( const ::rtl::OUString& rChars )
                            throw(  ::com::sun::star::xml::sax::SAXException,
                                    ::com::sun::star::uno::RuntimeException );

    virtual void SAL_CALL endFastElement( sal_Int32 nElement )
                            throw(  ::com::sun::star::xml::sax::SAXException,
                                    ::com::sun::star::uno::RuntimeException );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        onCreateContext( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

    // other ------------------------------------------------------------------

    /** Returns the full path of this fragment. */
    inline const ::rtl::OUString& getFragmentPath() const { return maFragmentPath; }

    /** Returns the relation with the passed type. */
    ::oox::core::RelationPtr getRelationByType( const ::rtl::OUString& rType ) const;
    /** Returns the relation with the passed identifier. */
    ::oox::core::RelationPtr getRelationById( const ::rtl::OUString& rRelId ) const;

    /** Returns the full path of the fragment from the passed relation. */
    ::rtl::OUString     getFragmentPathByRelation( const ::oox::core::Relation& rRelation ) const;
    /** Returns the full path of the fragment with the passed type. */
    ::rtl::OUString     getFragmentPathByType( const ::rtl::OUString& rType ) const;
    /** Returns the full path of the fragment with the passed relation identifier. */
    ::rtl::OUString     getFragmentPathByRelId( const ::rtl::OUString& rRelId ) const;

private:
    ::rtl::OUString     maFragmentPath;
};

// ============================================================================

template< typename HelperType >
class HelperFragmentBase : public ExcelFragmentBase, public HelperType
{
public:
    explicit            HelperFragmentBase(
                            const HelperType& rHelper,
                            const ::rtl::OUString& rFragmentPath );
};

template< typename HelperType >
HelperFragmentBase< HelperType >::HelperFragmentBase( const HelperType& rHelper, const ::rtl::OUString& rFragmentPath ) :
    ExcelFragmentBase( rHelper.getOoxFilter(), rFragmentPath ),
    HelperType( rHelper )
{
}

// ----------------------------------------------------------------------------

typedef HelperFragmentBase< GlobalDataHelper > GlobalFragmentBase;

typedef HelperFragmentBase< WorksheetHelper > WorksheetFragmentBase;

// ============================================================================

} // namespace xls
} // namespace oox

#endif

