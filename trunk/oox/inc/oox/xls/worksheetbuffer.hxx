/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: worksheetbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.1 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 12:29:40 $
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

#ifndef OOX_XLS_WORKSHEETBUFFER_HXX
#define OOX_XLS_WORKSHEETBUFFER_HXX

#include <vector>
#include <map>
#include "oox/xls/globaldatahelper.hxx"

namespace com { namespace sun { namespace star {
    namespace i18n { class XCharacterClassification; }
} } }

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

/** Contains data from the 'sheet' element describing a sheet in the workbook. */
struct OoxSheetInfo
{
    ::rtl::OUString     maId;           /// Relation identifier for the sheet substream.
    ::rtl::OUString     maName;         /// Original name of the sheet.
    ::rtl::OUString     maFinalName;    /// Final (converted) name of the sheet.
    bool                mbVisible;      /// True = sheet is visible.

    explicit            OoxSheetInfo();
};

// ============================================================================

/** Stores information about all internal and external linked sheets in a
    spreadsheet document.

    Information about sheets includes the sheet name, the visibility state, and
    for the OOX filter, the relation identifier of the sheet used to obtain the
    related worksheet fragment.
 */
class WorksheetBuffer : public GlobalDataHelper
{
public:
    explicit            WorksheetBuffer( const GlobalDataHelper& rGlobalData );

    /** Imports the attributes of a sheet element. */
    void                importSheet( const ::oox::core::AttributeList& rAttribs );
    /** Imports the BOUNDSHEET record from the passed BIFF stream. */
    void                importBoundSheet( BiffInputStream& rStrm );

    sal_Int32           insertExternalSheet(
                            const ::rtl::OUString& rTargetUrl,
                            const ::rtl::OUString& rSheetName );

    /** Returns the number of internal sheets. */
    sal_Int32           getInternalSheetCount() const;
    /** Returns the OOX relation identifier of the specified sheet. */
    ::rtl::OUString     getSheetRelId( sal_Int32 nSheet ) const;
    /** Returns the finalized name of the specified sheet. */
    ::rtl::OUString     getFinalSheetName( sal_Int32 nSheet ) const;
    /** Returns the finalized name of the sheet with the passed original name. */
    ::rtl::OUString     getFinalSheetName( const ::rtl::OUString& rName ) const;
    /** Returns the index of the sheet with the passed original name. */
    sal_Int32           getFinalSheetIndex( const ::rtl::OUString& rName ) const;

private:
    sal_Int16           getTotalSheetCount() const;
    const OoxSheetInfo* getSheetInfo( sal_Int32 nSheet ) const;
    ::rtl::OUString     convertToValidSheetName( const ::rtl::OUString& rName, sal_Unicode cReplaceChar ) const;

    ::rtl::OUString     insertSheet( const ::rtl::OUString& rName, sal_Int16 nSheet, bool bVisible );
    void                insertSheet( const OoxSheetInfo& rSheetInfo );

private:
    typedef ::std::vector< OoxSheetInfo >                   SheetInfoVec;
    typedef ::std::pair< ::rtl::OUString, ::rtl::OUString > ExternalSheetName;
    typedef ::std::map< ExternalSheetName, sal_Int16 >      ExternalSheetMap;

    ::com::sun::star::uno::Reference< ::com::sun::star::i18n::XCharacterClassification > mxCharClass;
    SheetInfoVec        maSheetInfos;
    ExternalSheetMap    maExternalSheets;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

