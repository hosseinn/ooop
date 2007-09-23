/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: externallinkbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.6 $
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

#ifndef OOX_XLS_EXTERNALLINKBUFFER_HXX
#define OOX_XLS_EXTERNALLINKBUFFER_HXX

#include "oox/core/containerhelper.hxx"
#include "oox/xls/globaldatahelper.hxx"

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

const sal_uInt16 BIFF_EXTERNNAME_BUILTIN    = 0x0001;
const sal_uInt16 BIFF_EXTERNNAME_AUTOMATIC  = 0x0002;
const sal_uInt16 BIFF_EXTERNNAME_PREFERPIC  = 0x0004;
const sal_uInt16 BIFF_EXTERNNAME_STDDOCNAME = 0x0008;
const sal_uInt16 BIFF_EXTERNNAME_OLEOBJECT  = 0x0010;
const sal_uInt16 BIFF_EXTERNNAME_ICONIFIED  = 0x8000;

// ============================================================================

/** Contains indexes for a range of sheets in the spreadsheet document. */
struct LinkSheetRange
{
    sal_Int32           mnFirst;        /// Index of the first sheet.
    sal_Int32           mnLast;         /// Index of the last sheet.

    inline explicit     LinkSheetRange() { setDeleted(); }
    inline explicit     LinkSheetRange( sal_Int32 nFirst, sal_Int32 nLast ) { set( nFirst, nLast ); }

    /** Sets this struct to deleted state. */
    inline void         setDeleted() { mnFirst = mnLast = -1; }
    /** Sets the passed sheet range to the memebers of this struct. */
    inline void         set( sal_Int32 nFirst, sal_Int32 nLast )
                            { mnFirst = ::std::min( nFirst, nLast ); mnLast = ::std::max( nFirst, nLast ); }

    /** Returns true, if one of the sheet indexes in invalid (negative). */
    inline bool         isDeleted() const { return (mnFirst < 0) || (mnLast < 0); }
    /** Returns true, if the sheet indexes are valid and different. */
    inline bool         is3dRange() const { return !isDeleted() && (mnFirst < mnLast); }
};

// ============================================================================

class ExternalName : public GlobalDataHelper
{
public:
    explicit            ExternalName( const GlobalDataHelper& rGlobalData );

    /** Imports the definedName element. */
    void                importDefinedName( const ::oox::core::AttributeList& rAttribs );
    /** Imports the ddeItem element describing an item of a DDE link. */
    void                importDdeItem( const ::oox::core::AttributeList& rAttribs );
    /** Imports the oleItem element describing an object of an OLE link. */
    void                importOleItem( const ::oox::core::AttributeList& rAttribs );

    /** Imports the EXTERNNAME record from the passed stream. */
    void                importExternName( BiffInputStream& rStrm );

    /** Returns the name of this external name. */
    inline const ::rtl::OUString& getName() const { return maName; }
    /** Returns true, if the name refers to an OLE object. */
    inline bool         isOleObject() const { return mbOleObj; }

private:
    ::rtl::OUString     maName;             /// Name of the external name.
    bool                mbBuiltIn;          /// Name is a built-in name.
    bool                mbStdDocName;       /// Name is the StdDocumentName for DDE.
    bool                mbOleObj;           /// Name is an OLE object.
    bool                mbNotify;           /// Notify application on data change.
    bool                mbPreferPic;        /// Picture link.
    bool                mbIconified;        /// Iconified object link.
};

typedef ::boost::shared_ptr< ExternalName > ExternalNameRef;

// ============================================================================

enum ExternalLinkType
{
    LINKTYPE_EXTERNAL,      /// Link refers to an external spreadsheet document.
    LINKTYPE_INTERNAL,      /// Link refers to own workbook.
    LINKTYPE_ANALYSIS,      /// Link refers to Analysis add-in.
    LINKTYPE_DDE,           /// DDE link.
    LINKTYPE_OLE,           /// OLE link.
    LINKTYPE_MAYBE_DDE_OLE, /// Could be DDE or OLE link (BIFF only).
    LINKTYPE_UNKNOWN        /// Unknown or unsupported link type.
};

// ============================================================================

class ExternalLink : public GlobalDataHelper
{
public:
    explicit            ExternalLink( const GlobalDataHelper& rGlobalData );

    /** Imports the externalBook element describing an externally linked document. */
    void                importExternalBook( const ::oox::core::AttributeList& rAttribs, const ::rtl::OUString& rTargetUrl );
    /** Imports the sheetName element containing the sheet name in an externally linked document. */
    void                importSheetName( const ::oox::core::AttributeList& rAttribs );
    /** Imports the definedName element describing an external name. */
    void                importDefinedName( const ::oox::core::AttributeList& rAttribs );
    /** Imports the ddeLink element describing a DDE link. */
    void                importDdeLink( const ::oox::core::AttributeList& rAttribs );
    /** Imports the ddeItem element describing an item of a DDE link. */
    void                importDdeItem( const ::oox::core::AttributeList& rAttribs );
    /** Imports the oleLink element describing an OLE link. */
    void                importOleLink( const ::oox::core::AttributeList& rAttribs, const ::rtl::OUString& rTargetUrl );
    /** Imports the oleItem element describing an object of an OLE link. */
    void                importOleItem( const ::oox::core::AttributeList& rAttribs );

    /** Imports the EXTERNSHEET record from the passed stream. */
    void                importExternSheet( BiffInputStream& rStrm );
    /** Imports the SUPBOOK record from the passed stream. */
    void                importSupBook( BiffInputStream& rStrm );
    /** Imports the EXTERNNAME record from the passed stream. */
    void                importExternName( BiffInputStream& rStrm );

    /** Returns the type of this external link. */
    inline ExternalLinkType getLinkType() const { return meLinkType; }
    /** Returns the class name of this external link. */
    inline const ::rtl::OUString& getClassName() const { return maClassName; }
    /** Returns the target URL of this external link. */
    inline const ::rtl::OUString& getTargetUrl() const { return maTargetUrl; }

    /** Returns the internal sheet index for the specified external sheet index. */
    sal_Int32           getSheetIndex( sal_Int32 nTabId ) const;
    /** Returns the sheet range for the specified sheets (BIFF only). */
    void                getSheetRange( LinkSheetRange& orSheetRange, sal_Int16 nTabId1, sal_Int16 nTabId2 ) const;
    /** Returns the external name with the passed one-based name index (BIFF only). */
    ExternalNameRef     getExternalName( sal_uInt16 nNameId ) const;

private:
    ExternalNameRef     createExternalName();
    ::rtl::OUString     parseBiffTarget( const ::rtl::OUString& rBiffEncoded );

private:
    typedef ::std::vector< sal_Int32 >                  SheetIndexVec;
    typedef ::oox::core::RefVector< ExternalName >      ExternalNameVec;

    ExternalLinkType    meLinkType;
    ::rtl::OUString     maClassName;
    ::rtl::OUString     maTargetUrl;
    SheetIndexVec       maSheetIndexes;
    ExternalNameVec     maExtNames;
};

typedef ::boost::shared_ptr< ExternalLink > ExternalLinkRef;

// ============================================================================

class ExternalLinkBuffer : public GlobalDataHelper
{
public:
    explicit            ExternalLinkBuffer( const GlobalDataHelper& rGlobalData );

    /** Imports the external reference data from the externalReference element. */
    ExternalLinkRef     importExternalReference( const ::oox::core::AttributeList& rAttribs );

    /** Imports the EXTERNSHEET record from the passed stream. */
    void                importExternSheet( BiffInputStream& rStrm );
    /** Imports the SUPBOOK record from the passed stream. */
    void                importSupBook( BiffInputStream& rStrm );
    /** Imports the EXTERNNAME record from the passed stream. */
    void                importExternName( BiffInputStream& rStrm );

    /** Returns the external link for the passed reference identifier. */
    ExternalLinkRef     getExternalLink( sal_Int32 nRefId ) const;
    /** Returns the link type for the passed reference identifier. */
    ExternalLinkType    getLinkType( sal_Int32 nRefId ) const;

    /** Returns the sheet range for the specified reference (BIFF5 only). */
    LinkSheetRange      getSheetRange( sal_Int32 nRefId, sal_Int16 nTabId1, sal_Int16 nTabId2 ) const;
    /** Returns the sheet range for the specified reference (BIFF8 only). */
    LinkSheetRange      getSheetRange( sal_Int32 nRefId ) const;
    /** Returns the external name with the passed one-based name index (BIFF only). */
    ExternalNameRef     getExternalName( sal_Int32 nRefId, sal_uInt16 nNameId ) const;

private:
    /** Represents a REF entry in the BIFF8 EXTERNSHEET record, maps ref id's to
        SUPBOOK records, and provides sheet indexes into the SUPBOOK sheet list. */
    struct Biff8RefEntry
    {
        sal_uInt16          mnSupBookId;        /// Zero-based index to SUPBOOK record.
        sal_Int16           mnTabId1;           /// Zero-based index to first sheet in SUPBOOK record.
        sal_Int16           mnTabId2;           /// Zero-based index to last sheet in SUPBOOK record.
    };

    /** Creates a new external link and inserts it into the list of links. */
    ExternalLinkRef     createExternalLink();

    /** Returns the specified BIFF8 REF entry from the EXTERNSHEET record. */
    const Biff8RefEntry* getRefEntry( sal_Int32 nRefId ) const;

private:
    typedef ::oox::core::RefVector< ExternalLink >  ExternalLinkVec;
    typedef ::std::vector< Biff8RefEntry >          Biff8RefEntryVec;

    ExternalLinkVec     maExtLinks;         /// List of external documents.
    Biff8RefEntryVec    maRefEntries;       /// Contents of BIFF8 EXTERNSHEET record.
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

