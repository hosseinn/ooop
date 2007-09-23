/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: externallinkfragment.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
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

#ifndef OOX_XLS_EXTERNALLINKFRAGMENT_HXX
#define OOX_XLS_EXTERNALLINKFRAGMENT_HXX

#include "oox/xls/excelfragmentbase.hxx"

namespace oox {
namespace xls {

// ============================================================================

class ExternalLink;

class OoxExternalLinkFragment : public GlobalFragmentBase
{
public:
    explicit            OoxExternalLinkFragment(
                            const GlobalDataHelper& rGlobalData,
                            ExternalLink& rExtLink,
                            const ::rtl::OUString& rFragmentPath );

    // oox.xls.ContextHelper interface ----------------------------------------

    virtual bool        onCanCreateContext( sal_Int32 nElement );
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        onCreateContext( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs );

private:
    ExternalLink&       mrExtLink;
    sal_Int32           mnCurrSheetId;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

