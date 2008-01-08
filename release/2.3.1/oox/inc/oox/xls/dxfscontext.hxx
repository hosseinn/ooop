/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: dxfscontext.hxx,v $
 *
 *  $Revision: 1.1.2.5 $
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

#ifndef OOX_XLS_DXFSCONTEXT_HXX
#define OOX_XLS_DXFSCONTEXT_HXX

#include "oox/xls/excelcontextbase.hxx"
#include <boost/shared_ptr.hpp>

namespace oox {
namespace xls {

// ============================================================================

class Dxf;

class OoxDxfsContext : public GlobalContextBase
{
public:
    explicit OoxDxfsContext( const GlobalFragmentBase& rFragment );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );
    virtual void        onEndElement( const ::rtl::OUString& rChars );

private:
    void                importColor( const ::oox::core::AttributeList& rAttribs );

private:
    ::boost::shared_ptr<Dxf>    mpCurDxf;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

