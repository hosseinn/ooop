/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: externallinkbuffer.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
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

#include "oox/xls/externallinkbuffer.hxx"
#include "oox/core/attributelist.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/worksheetbuffer.hxx"

using ::rtl::OString;
using ::rtl::OUString;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

// ============================================================================

ExternalName::ExternalName( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mbStdDocName( false ),
    mbOleObj( false ),
    mbNotify( false ),
    mbPreferPic( false ),
    mbIconified( false )
{
}

void ExternalName::importDefinedName( const AttributeList& rAttribs )
{
    maName = rAttribs.getString( XML_name );
    OSL_ENSURE( maName.getLength() > 0, "ExternalName::importDefinedName - empty name" );
}

void ExternalName::importDdeItem( const AttributeList& rAttribs )
{
    maName = rAttribs.getString( XML_name );
    OSL_ENSURE( maName.getLength() > 0, "ExternalName::importDdeItem - empty name" );
    mbStdDocName = rAttribs.getBool( XML_ole, false );
    mbNotify = rAttribs.getBool( XML_advise, false );
    mbPreferPic = rAttribs.getBool( XML_preferPic, false );
}

void ExternalName::importOleItem( const AttributeList& rAttribs )
{
    maName = rAttribs.getString( XML_name );
    OSL_ENSURE( maName.getLength() > 0, "ExternalName::importOleItem - empty name" );
    mbOleObj = true;
    mbNotify = rAttribs.getBool( XML_advise, false );
    mbPreferPic = rAttribs.getBool( XML_preferPic, false );
    mbIconified = rAttribs.getBool( XML_icon, false );
}

void ExternalName::importExternName( BiffInputStream& rStrm )
{
    sal_uInt16 nFlags = (getBiff() >= BIFF4) ? rStrm.readuInt16() : 0;
    mbBuiltIn = getFlag( nFlags, BIFF_EXTERNNAME_BUILTIN );

    if( getBiff() >= BIFF5 )
    {
        mbStdDocName = getFlag( nFlags, BIFF_EXTERNNAME_STDDOCNAME );
        mbOleObj     = getFlag( nFlags, BIFF_EXTERNNAME_OLEOBJECT );
        mbNotify     = getFlag( nFlags, BIFF_EXTERNNAME_AUTOMATIC );
        mbPreferPic  = getFlag( nFlags, BIFF_EXTERNNAME_PREFERPIC );
        mbIconified  = getFlag( nFlags, BIFF_EXTERNNAME_ICONIFIED );
        rStrm.ignore( 4 );
    }

    maName = (getBiff() == BIFF8) ?
        rStrm.readUniString( rStrm.readuInt8() ) :
        rStrm.readByteString( false, getTextEncoding() );
    OSL_ENSURE( maName.getLength() > 0, "ExternalName::importExternName - empty name" );
}

// ============================================================================

ExternalLink::ExternalLink( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    meLinkType( LINKTYPE_UNKNOWN )
{
}

void ExternalLink::importExternalBook( const AttributeList&, const OUString& rTargetUrl )
{
    maTargetUrl = rTargetUrl;
    OSL_ENSURE( rTargetUrl.getLength() > 0, "ExternalLink::importExternalBook - empty target URL" );
    meLinkType = (rTargetUrl.getLength() > 0) ? LINKTYPE_EXTERNAL : LINKTYPE_UNKNOWN;
}

void ExternalLink::importSheetName( const AttributeList& rAttribs )
{
    OUString aSheetName = rAttribs.getString( XML_val );
    OSL_ENSURE( aSheetName.getLength() > 0, "ExternalLink::importSheetName - empty sheet name" );
    if( meLinkType == LINKTYPE_EXTERNAL )
        maSheetIndexes.push_back( getWorksheets().insertExternalSheet( maTargetUrl, aSheetName ) );
}

void ExternalLink::importDefinedName( const AttributeList& rAttribs )
{
    createExternalName()->importDefinedName( rAttribs );
}

void ExternalLink::importDdeLink( const AttributeList& rAttribs )
{
    maClassName = rAttribs.getString( XML_ddeService );
    OSL_ENSURE( maClassName.getLength() > 0, "ExternalLink::importDdeLink - empty DDE service name" );
    maTargetUrl = rAttribs.getString( XML_ddeTopic );
    OSL_ENSURE( maTargetUrl.getLength() > 0, "ExternalLink::importDdeLink - empty DDE topic" );
    meLinkType = ((maClassName.getLength() > 0) && (maTargetUrl.getLength() > 0)) ? LINKTYPE_DDE : LINKTYPE_UNKNOWN;
}

void ExternalLink::importDdeItem( const AttributeList& rAttribs )
{
    createExternalName()->importDdeItem( rAttribs );
}

void ExternalLink::importOleLink( const AttributeList& rAttribs, const OUString& rTargetUrl )
{
    maClassName = rAttribs.getString( XML_progId );
    OSL_ENSURE( maClassName.getLength() > 0, "ExternalLink::importOleLink - empty OLE class name" );
    maTargetUrl = rTargetUrl;
    OSL_ENSURE( rTargetUrl.getLength() > 0, "ExternalLink::importOleLink - empty target URL" );
    meLinkType = ((maClassName.getLength() > 0) && (rTargetUrl.getLength() > 0)) ? LINKTYPE_OLE : LINKTYPE_UNKNOWN;
}

void ExternalLink::importOleItem( const AttributeList& rAttribs )
{
    createExternalName()->importOleItem( rAttribs );
}

void ExternalLink::importExternSheet( BiffInputStream& rStrm )
{
    OString aTarget = rStrm.readByteString( false );
    // references to own sheets have wrong string length field (off by 1)
    if( (aTarget.getLength() > 0) && (aTarget[ 0 ] == 3) )
        aTarget += OString( static_cast< sal_Char >( rStrm.readuInt8() ) );
    // parse the encoded URL
    OUString aSheetName = parseBiffTarget( OStringToOUString( aTarget, getTextEncoding() ) );
    // create the linked sheet in the Calc document
    if( (meLinkType == LINKTYPE_EXTERNAL) && (aSheetName.getLength() > 0) )
        maSheetIndexes.push_back( getWorksheets().insertExternalSheet( maTargetUrl, aSheetName ) );
}

void ExternalLink::importSupBook( BiffInputStream& rStrm )
{
    OUString aTarget;
    sal_uInt16 nSheetCount;
    rStrm >> nSheetCount;
    if( rStrm.getRecLeft() == 2 )
    {
        if( rStrm.readuInt8() == 1 )
        {
            sal_Char cChar = static_cast< sal_Char >( rStrm.readuInt8() );
            if( cChar != 0 )
                aTarget = OStringToOUString( OString( cChar ), getTextEncoding() );
        }
    }
    else if( rStrm.getRecLeft() >= 3 )
    {
        aTarget = rStrm.readUniString();
    }

    // parse the encoded URL
    OUString aSheetName = parseBiffTarget( aTarget );
    OSL_ENSURE( aSheetName.getLength() == 0, "ExternalLink::importSupBook - sheet name in encoded URL" );
    (void)aSheetName;   // prevent compiler warning

    // load external sheet names and create the linked sheets in the Calc document
    if( meLinkType == LINKTYPE_EXTERNAL )
    {
        WorksheetBuffer& rWorksheets = getWorksheets();
        for( sal_uInt16 nSheet = 0; rStrm.isValid() && (nSheet < nSheetCount); ++nSheet )
        {
            OUString aSheetName = rStrm.readUniString();
            OSL_ENSURE( aSheetName.getLength() > 0, "ExternalLink::importSupBook - empty sheet name" );
            maSheetIndexes.push_back( rWorksheets.insertExternalSheet( maTargetUrl, aSheetName ) );
        }
    }
}

void ExternalLink::importExternName( BiffInputStream& rStrm )
{
    ExternalNameRef xExtName = createExternalName();
    xExtName->importExternName( rStrm );
    switch( meLinkType )
    {
        case LINKTYPE_DDE:
            OSL_ENSURE( !xExtName->isOleObject(), "ExternalLink::importExternName - OLE object in DDE link" );
        break;
        case LINKTYPE_OLE:
            OSL_ENSURE( xExtName->isOleObject(), "ExternalLink::importExternName - anything but OLE object in OLE link" );
        break;
        case LINKTYPE_MAYBE_DDE_OLE:
            meLinkType = xExtName->isOleObject() ? LINKTYPE_OLE : LINKTYPE_DDE;
        break;
        default:
            OSL_ENSURE( !xExtName->isOleObject(), "ExternalLink::importExternName - OLE object in external name" );
    }
}

sal_Int32 ExternalLink::getSheetIndex( sal_Int32 nTabId ) const
{
    return ((0 <= nTabId) && (static_cast< size_t >( nTabId ) < maSheetIndexes.size())) ?
        maSheetIndexes[ static_cast< size_t >( nTabId ) ] : -1;
}

void ExternalLink::getSheetRange( LinkSheetRange& orSheetRange, sal_Int16 nTabId1, sal_Int16 nTabId2 ) const
{
    switch( meLinkType )
    {
        case LINKTYPE_INTERNAL:
            orSheetRange.set( nTabId1, nTabId2 );
        break;
        case LINKTYPE_EXTERNAL:
            orSheetRange.set( getSheetIndex( nTabId1 ), getSheetIndex( nTabId2 ) );
        break;
        default:
            orSheetRange.setDeleted();
    }
}

ExternalNameRef ExternalLink::getExternalName( sal_uInt16 nNameId ) const
{
    OSL_ENSURE( nNameId > 0, "ExternalLink::getExternalName - invalid name index" );
    return maExtNames.get( static_cast< sal_Int32 >( nNameId ) - 1 );
}

// private --------------------------------------------------------------------

ExternalNameRef ExternalLink::createExternalName()
{
    ExternalNameRef xExtName( new ExternalName( getGlobalData() ) );
    maExtNames.push_back( xExtName );
    return xExtName;
}

OUString ExternalLink::parseBiffTarget( const OUString& rBiffEncoded )
{
    OUString aSheetName;
    if( getAddressConverter().parseEncodedTarget( maClassName, maTargetUrl, aSheetName, rBiffEncoded ) )
    {
        if( maClassName.getLength() > 0 )
        {
            if( maTargetUrl.getLength() > 0 )
                meLinkType = LINKTYPE_MAYBE_DDE_OLE;
        }
        else if( maTargetUrl.getLength() == 0 )
        {
            meLinkType = LINKTYPE_INTERNAL;
        }
        else if( (maTargetUrl.getLength() == 1) && (maTargetUrl[ 0 ] == ':') )
        {
            if( getBiff() >= BIFF4 )
                meLinkType = LINKTYPE_ANALYSIS;
        }
        else if( (maTargetUrl.getLength() == 1) && (maTargetUrl[ 0 ] == ' ') )
        {
            meLinkType = LINKTYPE_UNKNOWN;
        }
        else
        {
            meLinkType = LINKTYPE_EXTERNAL;
        }
    }
    else
    {
        meLinkType = LINKTYPE_UNKNOWN;
    }
    return aSheetName;
}

// ============================================================================

ExternalLinkBuffer::ExternalLinkBuffer( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

ExternalLinkRef ExternalLinkBuffer::importExternalReference( const AttributeList& )
{
    return createExternalLink();
}

void ExternalLinkBuffer::importExternSheet( BiffInputStream& rStrm )
{
    switch( getBiff() )
    {
        case BIFF2:
        case BIFF3:
        case BIFF4:
        case BIFF5:
            createExternalLink()->importExternSheet( rStrm );
        break;
        case BIFF8:
        {
            OSL_ENSURE( maRefEntries.empty(), "ExternalLinkBuffer::importExternSheet - multiple EXTERNSHEET records" );
            sal_uInt16 nRefCount;
            rStrm >> nRefCount;
            maRefEntries.reserve( nRefCount );
            for( sal_uInt16 nRefId = 0; rStrm.isValid() && (nRefId < nRefCount); ++nRefId )
            {
                Biff8RefEntry aRefEntry;
                rStrm >> aRefEntry.mnSupBookId >> aRefEntry.mnTabId1 >> aRefEntry.mnTabId2;
                maRefEntries.push_back( aRefEntry );
            }
        }
        break;
        case BIFF_UNKNOWN: break;
    }
}

void ExternalLinkBuffer::importSupBook( BiffInputStream& rStrm )
{
    createExternalLink()->importSupBook( rStrm );
}

void ExternalLinkBuffer::importExternName( BiffInputStream& rStrm )
{
    if( !maExtLinks.empty() )
        maExtLinks.back()->importExternName( rStrm );
}

ExternalLinkRef ExternalLinkBuffer::getExternalLink( sal_Int32 nRefId ) const
{
    ExternalLinkRef xExtLink;
    switch( getFilterType() )
    {
        case FILTER_OOX:
            // one-based index
            xExtLink = maExtLinks.get( nRefId - 1 );
        break;
        case FILTER_BIFF:
            switch( getBiff() )
            {
                case BIFF2:
                case BIFF3:
                case BIFF4:
                break;
                case BIFF5:
                    // one-based index to EXTERNSHEET records
                    if( nRefId < 0 )
                        xExtLink = maExtLinks.get( 1 - nRefId );
                    else if( nRefId > 0 )
                        xExtLink = maExtLinks.get( nRefId - 1 );
                    // internal links have negative index
                    if( xExtLink.get() && ((nRefId < 0) != (xExtLink->getLinkType() == LINKTYPE_INTERNAL)) )
                        xExtLink.reset();
                break;
                case BIFF8:
                    // zero-based index into REF list in EXTERNSHEET record
                    if( const Biff8RefEntry* pRefEntry = getRefEntry( nRefId ) )
                        xExtLink = maExtLinks.get( pRefEntry->mnSupBookId );
                break;
                case BIFF_UNKNOWN: break;
            }
        break;
        case FILTER_UNKNOWN: break;
    }
    return xExtLink;
}

ExternalLinkType ExternalLinkBuffer::getLinkType( sal_Int32 nRefId ) const
{
    ExternalLinkRef xExtLink = getExternalLink( nRefId );
    return xExtLink.get() ? xExtLink->getLinkType() : LINKTYPE_UNKNOWN;
}

LinkSheetRange ExternalLinkBuffer::getSheetRange( sal_Int32 nRefId, sal_Int16 nTabId1, sal_Int16 nTabId2 ) const
{
    OSL_ENSURE( getBiff() == BIFF5, "ExternalLinkBuffer::getSheetRange - wrong BIFF version" );
    LinkSheetRange aSheetRange;
    if( const ExternalLink* pExtLink = getExternalLink( nRefId ).get() )
        pExtLink->getSheetRange( aSheetRange, nTabId1, nTabId2 );
    return aSheetRange;
}

LinkSheetRange ExternalLinkBuffer::getSheetRange( sal_Int32 nRefId ) const
{
    OSL_ENSURE( getBiff() == BIFF8, "ExternalLinkBuffer::getSheetRange - wrong BIFF version" );
    LinkSheetRange aSheetRange;
    if( const ExternalLink* pExtLink = getExternalLink( nRefId ).get() )
        if( const Biff8RefEntry* pRefEntry = getRefEntry( nRefId ) )
            pExtLink->getSheetRange( aSheetRange, pRefEntry->mnTabId1, pRefEntry->mnTabId2 );
    return aSheetRange;
}

ExternalNameRef ExternalLinkBuffer::getExternalName( sal_Int32 nRefId, sal_uInt16 nNameId ) const
{
    ExternalNameRef xExtName;
    if( const ExternalLink* pExtLink = getExternalLink( nRefId ).get() )
        xExtName = pExtLink->getExternalName( nNameId );
    return xExtName;
}

// private --------------------------------------------------------------------

ExternalLinkRef ExternalLinkBuffer::createExternalLink()
{
    ExternalLinkRef xExtLink( new ExternalLink( getGlobalData() ) );
    maExtLinks.push_back( xExtLink );
    return xExtLink;
}

const ExternalLinkBuffer::Biff8RefEntry* ExternalLinkBuffer::getRefEntry( sal_Int32 nRefId ) const
{
    return ((0 <= nRefId) && (static_cast< size_t >( nRefId ) < maRefEntries.size())) ?
        &maRefEntries[ static_cast< size_t >( nRefId ) ] : 0;
}

// ============================================================================

} // namespace xls
} // namespace oox

