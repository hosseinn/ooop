/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: connectionsfragment.hxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/01 09:58:26 $
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

#ifndef OOX_XLS_CONNECTIONSFRAGMENT_HXX
#define OOX_XLS_CONNECTIONSFRAGMENT_HXX

#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/excelfragmentbase.hxx"
#include "rtl/ustring.hxx"

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

class OoxConnectionsFragment : public GlobalFragmentBase
{
public:
    explicit            OoxConnectionsFragment(
                            const GlobalDataHelper& rGlobalData,
                            const ::rtl::OUString& rFragmentPath );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );

private:
    void                importConnection( const ::oox::core::AttributeList& rAttribs );
    void                importWebPr( const ::oox::core::AttributeList& rAttribs );
    void                importTables( const ::oox::core::AttributeList& rAttribs );
    void                importS( const ::oox::core::AttributeList& rAttribs );
    void                importX( const ::oox::core::AttributeList& rAttribs );
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

