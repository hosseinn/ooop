/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sharedstringsbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.10 $
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

#ifndef OOX_XLS_SHAREDSTRINGSBUFFER_HXX
#define OOX_XLS_SHAREDSTRINGSBUFFER_HXX

#include "oox/core/containerhelper.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/richstring.hxx"

namespace oox {
namespace xls {

// ============================================================================

/** Collects all strings from the shared strings substream. */
class SharedStringsBuffer : public GlobalDataHelper
{
public:
    explicit            SharedStringsBuffer( const GlobalDataHelper& rGlobalData );

    /** Creates and returns a new string entry. */
    RichStringRef       createRichString();
    /** Imports the complete shared string table from a BIFF file. */
    void                importSst( BiffInputStream& rStrm );

    /** Final processing after import of all strings. */
    void                finalizeImport();

    /** Converts the specified string table entry. */
    void                convertString(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::text::XText >& rxText,
                            sal_Int32 nIndex,
                            sal_Int32 nXfId ) const;

private:
    typedef ::oox::core::RefVector< RichString > StringVec;
    StringVec           maStrings;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

