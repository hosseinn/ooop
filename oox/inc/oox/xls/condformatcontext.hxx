/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: condformatcontext.hxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:00 $
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

#ifndef OOX_XLS_CONDFORMATCONTEXT_HXX
#define OOX_XLS_CONDFORMATCONTEXT_HXX

#include "oox/xls/excelcontextbase.hxx"
#include "oox/xls/worksheethelper.hxx"
#include "oox/xls/condformatbuffer.hxx"

#include <com/sun/star/table/CellRangeAddress.hpp>
#include <rtl/ustrbuf.hxx>
#include <boost/shared_ptr.hpp>
#include <vector>

namespace oox {
namespace xls {

// ============================================================================

class OoxCondFormatContext : public WorksheetContextBase
{
public:
    explicit            OoxCondFormatContext( const WorksheetFragmentBase& rFragment );

    virtual void SAL_CALL startDocument() throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
    virtual void SAL_CALL endDocument() throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    void                importConditionalFormatting( const ::oox::core::AttributeList& rAttribs );
    void                importCfRule( const ::oox::core::AttributeList& rAttribs );
    void                importFormula( const ::oox::core::AttributeList& rAttribs );

private:

    CondFormatItemRef   mpCurItem;
    sal_Int16                                   mnEntryId;
    bool mbValidRange;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

