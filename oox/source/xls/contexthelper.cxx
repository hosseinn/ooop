/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: contexthelper.cxx,v $
 *
 *  $Revision: 1.1.2.11 $
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

#include "oox/xls/contexthelper.hxx"
#include <osl/diagnose.h>

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::com::sun::star::xml::sax::XFastAttributeList;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

// ============================================================================

ContextHelper::ContextHelper()
{
}

ContextHelper::~ContextHelper()
{
}

sal_Int32 ContextHelper::getPreviousContext( sal_Int32 nCountBack ) const
{
    if( (nCountBack < 0) || (maContextStack.size() < static_cast< size_t >( nCountBack )) )
        return XML_TOKEN_INVALID;
    return (maContextStack.size() == static_cast< size_t >( nCountBack )) ?
        XML_ROOT_CONTEXT : maContextStack[ maContextStack.size() - nCountBack - 1 ].mnElement;
}

Reference< XFastContextHandler > ContextHelper::implCreateChildContext( sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs )
{
    if( !maContextStack.empty() )
        appendCollectedChars();
    Reference< XFastContextHandler > xContext;
    if( onCanCreateContext( nElement ) )
        xContext = onCreateContext( nElement, AttributeList( rxAttribs ) );
    return xContext;
}

void ContextHelper::implStartCurrentContext( sal_Int32 nElement, const Reference< XFastAttributeList >& rxAttribs )
{
    AttributeList aAttribs( rxAttribs );
    maContextStack.resize( maContextStack.size() + 1 );
    ContextInfo& rInfo = maContextStack.back();
    rInfo.mnElement = nElement;
    rInfo.mbTrimSpaces = rxAttribs->getOptionalValueToken( XML_TOKEN( space ), XML_TOKEN_INVALID ) != XML_preserve;
    onStartElement( aAttribs );
}

void ContextHelper::implCharacters( const OUString& rChars )
{
    // #i76091# collect characters until context ends
    if( !maContextStack.empty() )
        maContextStack.back().maCurrChars.append( rChars );
}

void ContextHelper::implEndCurrentContext( sal_Int32 nElement )
{
    (void)nElement;     // prevent "unused parameter" warning in product build
    OSL_ENSURE( getCurrentContext() == nElement, "ContextHelper::implEndCurrentContext - context stack broken" );
    if( !maContextStack.empty() )
    {
        // #i76091# process collected characters
        appendCollectedChars();
        // finalize the current context and pop context info from stack
        onEndElement( maContextStack.back().maFinalChars.makeStringAndClear() );
        maContextStack.pop_back();
    }
}

void ContextHelper::appendCollectedChars()
{
    OSL_ENSURE( !maContextStack.empty(), "ContextHelper::appendCollectedChars - no context info" );
    ContextInfo& rInfo = maContextStack.back();
    if( rInfo.maCurrChars.getLength() > 0 )
    {
        OUString aChars = rInfo.maCurrChars.makeStringAndClear();
        rInfo.maFinalChars.append( rInfo.mbTrimSpaces ? aChars.trim() : aChars );
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

