/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: pivottablefragment.hxx,v $
 *
 *  $Revision: 1.1.2.8 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 12:48:01 $
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

#ifndef OOX_XLS_PIVOTTABLEFRAGMENT_HXX
#define OOX_XLS_PIVOTTABLEFRAGMENT_HXX

#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/worksheethelper.hxx"
#include "oox/xls/excelfragmentbase.hxx"
#include "oox/xls/pivottablebuffer.hxx"
#include "rtl/ustring.hxx"

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

class OoxPivotTableFragment : public WorksheetFragmentBase
{
public:
    explicit            OoxPivotTableFragment(
                            const WorksheetHelper& rSheetHelper,
                            const ::rtl::OUString& rFragmentPath );

    virtual void SAL_CALL endDocument() throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:

    void                importLocation( const ::oox::core::AttributeList& rAttribs );

    void                importPivotTableDefinition( const ::oox::core::AttributeList& rAttribs );

    void                importPivotFields( const ::oox::core::AttributeList& rAttribs );

    void                importPivotField( const ::oox::core::AttributeList& rAttribs );

private:
    ::rtl::OUString maName;
    PivotTableData  maData;

    bool        mbValidRange;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

