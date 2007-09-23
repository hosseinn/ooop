/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: richstring.cxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/31 11:40:04 $
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

#include "oox/xls/richstring.hxx"
#include <com/sun/star/text/XText.hpp>
#include "oox/core/attributelist.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/biffinputstream.hxx"

using ::rtl::OString;
using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::text::XText;
using ::com::sun::star::text::XTextRange;
using ::oox::core::AttributeList;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

RichStringPortion::RichStringPortion( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mnFontId( -1 )
{
}

void RichStringPortion::setText( const OUString& rText )
{
    maText = rText;
}

FontRef RichStringPortion::importFont( const AttributeList& )
{
    mxFont.reset( new Font( getGlobalData() ) );
    return mxFont;
}

void RichStringPortion::setFontId( sal_Int32 nFontId )
{
    mnFontId = nFontId;
}

void RichStringPortion::finalizeImport()
{
    if( mxFont.get() )
        mxFont->finalizeImport();
    else if( mnFontId >= 0 )
        mxFont = getStyles().getFont( mnFontId );
}

void RichStringPortion::convert( const Reference< XText >& rxText, sal_Int32 nXfId )
{
    Reference< XTextRange > xRange = rxText->getEnd();
    xRange->setString( maText );
    if( mxFont.get() )
    {
        PropertySet aPropSet( xRange );
        mxFont->writeToPropertySet( aPropSet, FONT_PROPTYPE_RICHTEXT );
    }
    if( const Font* pFont = getStyles().getFontFromCellXf( nXfId ).get() )
    {
        if( pFont->needsRichTextFormat() )
        {
            PropertySet aPropSet( xRange );
            pFont->writeToPropertySet( aPropSet, FONT_PROPTYPE_RICHTEXT );
        }
    }
}

// ============================================================================

RichStringPhonetic::RichStringPhonetic( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mnBasePos( -1 ),
    mnBaseEnd( -1 )
{
}

void RichStringPhonetic::setText( const OUString& rText )
{
    maText = rText;
}

void RichStringPhonetic::importPhoneticRun( const AttributeList& rAttribs )
{
    mnBasePos = rAttribs.getInteger( XML_sb, -1 );
    mnBaseEnd = rAttribs.getInteger( XML_eb, -1 );
}

void RichStringPhonetic::setBaseRange( sal_Int32 nBasePos, sal_Int32 nBaseEnd )
{
    mnBasePos = nBasePos;
    mnBaseEnd = nBaseEnd;
}

// ============================================================================

PhoneticSettings::PhoneticSettings( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mnType( XML_fullwidthKatakana ),
    mnAlignment( XML_left ),
    mnFontId( -1 )
{
}

void PhoneticSettings::importPhoneticPr( const AttributeList& rAttribs )
{
    mnType      = rAttribs.getToken( XML_type, XML_fullwidthKatakana );
    mnAlignment = rAttribs.getToken( XML_alignment, XML_left );
    mnFontId    = rAttribs.getInteger( XML_fontId, -1 );
}

void PhoneticSettings::setBiffData( sal_uInt16 nFlags, sal_uInt16 nFontId )
{
    switch( extractValue< sal_uInt16 >( nFlags, 0, 2 ) )
    {
        case BIFF_PHONETIC_TYPE_KATAKANA_HALF:  mnType = XML_halfwidthKatakana; break;
        case BIFF_PHONETIC_TYPE_KATAKANA_FULL:  mnType = XML_fullwidthKatakana; break;
        case BIFF_PHONETIC_TYPE_HIRAGANA:       mnType = XML_hiragana;          break;
        default:                                mnType = XML_halfwidthKatakana;
    }
    switch( extractValue< sal_uInt16 >( nFlags, 2, 2 ) )
    {
        case BIFF_PHONETIC_ALIGN_NONE:          mnAlignment = XML_noControl;    break;
        case BIFF_PHONETIC_ALIGN_LEFT:          mnAlignment = XML_left;         break;
        case BIFF_PHONETIC_ALIGN_CENTER:        mnAlignment = XML_center;       break;
        case BIFF_PHONETIC_ALIGN_DISTRIBUTED:   mnAlignment = XML_distributed;  break;
    }
    mnFontId = nFontId;
}

// ============================================================================

RichString::RichString( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maPhoneticSett( rGlobalData )
{
}

RichStringPortionRef RichString::importText( const AttributeList& )
{
    return createPortion();
}

RichStringPortionRef RichString::importRun( const AttributeList& )
{
    return createPortion();
}

RichStringPhoneticRef RichString::importPhoneticRun( const AttributeList& rAttribs )
{
    RichStringPhoneticRef xPhonetic = createPhonetic();
    xPhonetic->importPhoneticRun( rAttribs );
    return xPhonetic;
}

void RichString::importPhoneticPr( const AttributeList& rAttribs )
{
    maPhoneticSett.importPhoneticPr( rAttribs );
}

void RichString::appendFontId( BiffRichStringFontIdVec& orFontIds, sal_Int32 nPos, sal_Int32 nFontId )
{
    // #i33341# real life -- same character index may occur several times
    OSL_ENSURE( orFontIds.empty() || (orFontIds.back().mnPos <= nPos), "RichString::appendFontId - wrong char order" );
    if( orFontIds.empty() || (orFontIds.back().mnPos < nPos) )
        orFontIds.push_back( BiffRichStringFontId( nPos, nFontId ) );
    else
        orFontIds.back().mnFontId = nFontId;
}

void RichString::importFontIds( BiffRichStringFontIdVec& orFontIds, BiffInputStream& rStrm, sal_uInt16 nCount, bool b16Bit )
{
    orFontIds.clear();
    orFontIds.reserve( nCount );
    /*  #i33341# real life -- same character index may occur several times
        -> use appendFontId() to validate string position. */
    if( b16Bit )
    {
        for( sal_uInt16 nIndex = 0; rStrm.isValid() && (nIndex < nCount); ++nIndex )
        {
            sal_uInt16 nPos, nFontId;
            rStrm >> nPos >> nFontId;
            appendFontId( orFontIds, nPos, nFontId );
        }
    }
    else
    {
        for( sal_uInt16 nIndex = 0; rStrm.isValid() && (nIndex < nCount); ++nIndex )
        {
            sal_uInt8 nPos, nFontId;
            rStrm >> nPos >> nFontId;
            appendFontId( orFontIds, nPos, nFontId );
        }
    }
}

void RichString::importFontIds( BiffRichStringFontIdVec& orFontIds, BiffInputStream& rStrm, bool b16Bit )
{
    sal_uInt16 nCount = b16Bit ? rStrm.readuInt16() : rStrm.readuInt8();
    importFontIds( orFontIds, rStrm, nCount, b16Bit );
}

void RichString::importByteString( BiffInputStream& rStrm, rtl_TextEncoding eDefaultTextEnc, BiffStringFlags nFlags )
{
    OSL_ENSURE( !getFlag( nFlags, BIFF_STR_KEEPFONTS ), "RichString::importString - keep fonts not implemented" );
    OSL_ENSURE( !getFlag( nFlags, static_cast< BiffStringFlags >( ~(BIFF_STR_8BITLENGTH | BIFF_STR_EXTRAFONTS) ) ), "RichString::importByteString - unknown flag" );
    bool b8BitLength = getFlag( nFlags, BIFF_STR_8BITLENGTH );

    OString aBaseText = rStrm.readByteString( !b8BitLength );
    BiffRichStringFontIdVec aFontIds;
    if( getFlag( nFlags, BIFF_STR_EXTRAFONTS ) )
        importFontIds( aFontIds, rStrm, false );

    createBiffPortions( aBaseText, eDefaultTextEnc, aFontIds );
}

void RichString::importUniString( BiffInputStream& rStrm, BiffStringFlags nFlags )
{
    OSL_ENSURE( !getFlag( nFlags, BIFF_STR_KEEPFONTS ), "RichString::importUniString - keep fonts not implemented" );
    OSL_ENSURE( !getFlag( nFlags, static_cast< BiffStringFlags >( ~(BIFF_STR_8BITLENGTH | BIFF_STR_SMARTFLAGS) ) ), "RichString::importUniString - unknown flag" );
    bool b8BitLength = getFlag( nFlags, BIFF_STR_8BITLENGTH );

    // --- string header ---
    sal_uInt16 nChars = b8BitLength ? rStrm.readuInt8() : rStrm.readuInt16();
    sal_uInt8 nFlagField = 0;
    if( (nChars > 0) || !getFlag( nFlags, BIFF_STR_SMARTFLAGS ) )
        rStrm >> nFlagField;

    bool b16Bit, bFonts, bPhonetic;
    sal_uInt16 nFontCount;
    sal_uInt32 nPhoneticSize;
    rStrm.readExtendedUniStringHeader( b16Bit, bFonts, bPhonetic, nFontCount, nPhoneticSize, nFlagField );

    // --- character array ---
    OUString aBaseText = rStrm.readRawUniString( nChars, b16Bit );

    // --- formatting ---
    // #122185# bRich flag may be set, but format runs may be missing
    BiffRichStringFontIdVec aFontIds;
    if( rStrm.isValid() && (nFontCount > 0) )
        importFontIds( aFontIds, rStrm, nFontCount, true );

    // --- Asian phonetic information ---
    // #122185# bPhonetic flag may be set, but phonetic info may be missing
    OUString aPhoneticText;
    BiffPhoneticPortionVec aPhonetics;
    if( rStrm.isValid() && (nPhoneticSize > 0) )
        aPhoneticText = importPhonetic( aPhonetics, rStrm, nPhoneticSize );

    // build string portions and phonetic portions
    createBiffPortions( aBaseText, aFontIds );
    createBiffPhonetics( aPhoneticText, aPhonetics, aBaseText.getLength() );
}

void RichString::finalizeImport()
{
    maPortions.forEachMem( &RichStringPortion::finalizeImport );
}

void RichString::convert( const Reference< XText >& rxText, sal_Int32 nXfId ) const
{
    for( PortionVec::const_iterator aIt = maPortions.begin(), aEnd = maPortions.end(); aIt != aEnd; ++aIt )
    {
        (*aIt)->convert( rxText, nXfId );
        nXfId = -1;
    }
}

// private --------------------------------------------------------------------

RichStringPortionRef RichString::createPortion()
{
    RichStringPortionRef xPortion( new RichStringPortion( getGlobalData() ) );
    maPortions.push_back( xPortion );
    return xPortion;
}

RichStringPhoneticRef RichString::createPhonetic()
{
    RichStringPhoneticRef xPhonetic( new RichStringPhonetic( getGlobalData() ) );
    maPhonetics.push_back( xPhonetic );
    return xPhonetic;
}

void RichString::appendPhoneticPortion( BiffPhoneticPortionVec& orPhonetics, sal_Int32 nPos, sal_Int32 nBasePos, sal_Int32 nBaseLen )
{
    // same character index may occur several times
    OSL_ENSURE( orPhonetics.empty() || ((orPhonetics.back().mnPos <= nPos) &&
        (orPhonetics.back().mnBasePos + orPhonetics.back().mnBaseLen <= nBasePos)),
        "RichString::appendPhoneticPortion - wrong char order" );
    if( orPhonetics.empty() || (orPhonetics.back().mnPos < nPos) )
    {
        orPhonetics.push_back( BiffPhoneticPortion( nPos, nBasePos, nBaseLen ) );
    }
    else if( orPhonetics.back().mnPos == nPos )
    {
        orPhonetics.back().mnBasePos = nBasePos;
        orPhonetics.back().mnBaseLen = nBaseLen;
    }
}

void RichString::importPhoneticPortions( BiffPhoneticPortionVec& orPhonetics, BiffInputStream& rStrm, sal_uInt16 nCount )
{
    orPhonetics.clear();
    orPhonetics.reserve( nCount );
    for( sal_uInt16 nPortion = 0; nPortion < nCount; ++nPortion )
    {
        sal_uInt16 nPos, nBasePos, nBaseLen;
        rStrm >> nPos >> nBasePos >> nBaseLen;
        appendPhoneticPortion( orPhonetics, nPos, nBasePos, nBaseLen );
    }
}

OUString RichString::importPhonetic( BiffPhoneticPortionVec& orPhonetics, BiffInputStream& rStrm, sal_uInt32 nPhoneticSize )
{
    OUString aPhoneticText;
    if( rStrm.isValid() && (nPhoneticSize > 0) )
    {
        sal_uInt32 nPhoneticEnd = rStrm.getRecPos() + nPhoneticSize;
        OSL_ENSURE( nPhoneticSize > 14, "RichString::importPhonetic - wrong size of phonetic data" );
        if( nPhoneticSize > 14 )
        {
            sal_uInt16 nId, nSize, nFontId, nFlags, nPortionCount, nTextLen1, nTextLen2;
            rStrm >> nId >> nSize >> nFontId >> nFlags >> nPortionCount >> nTextLen1 >> nTextLen2;
            OSL_ENSURE( nId == 1, "RichString::importPhonetic - unknown phonetic data identifier" );
            sal_uInt32 nMinSize = static_cast< sal_uInt32 >( nSize + 4 );
            OSL_ENSURE( nMinSize <= nPhoneticSize, "RichString::importPhonetic - wrong size of phonetic data" );
            OSL_ENSURE( nTextLen1 == nTextLen2, "RichString::importPhonetic - wrong phonetic text length" );
            if( (nId == 1) && (nMinSize <= nPhoneticSize) && (nTextLen1 == nTextLen2) )
            {
                maPhoneticSett.setBiffData( nFlags, nFontId );
                if( nTextLen1 > 0 )
                {
                    nMinSize = static_cast< sal_uInt32 >( 2 * nTextLen1 + 6 * nPortionCount + 14 );
                    OSL_ENSURE( nMinSize <= nPhoneticSize, "RichString::importUniString - wrong size of phonetic data" );
                    if( nMinSize <= nPhoneticSize )
                    {
                        aPhoneticText = rStrm.readUnicodeArray( nTextLen1 );
                        importPhoneticPortions( orPhonetics, rStrm, nPortionCount );
                    }
                }
            }
        }
        rStrm.seek( nPhoneticEnd );
    }
    return aPhoneticText;
}

void RichString::createBiffPortions( const OString& rText, rtl_TextEncoding eDefaultTextEnc, BiffRichStringFontIdVec& rFontIds )
{
    maPortions.clear();
    sal_Int32 nStrLen = rText.getLength();
    if( nStrLen > 0 )
    {
        // add leading and trailing string position to ease the following loop
        if( rFontIds.empty() || (rFontIds.front().mnPos > 0) )
            rFontIds.insert( rFontIds.begin(), BiffRichStringFontId( 0, -1 ) );
        if( rFontIds.back().mnPos < nStrLen )
            rFontIds.push_back( BiffRichStringFontId( nStrLen, -1 ) );

        // create all string portions according to the font id vector
        for( BiffRichStringFontIdVec::const_iterator aIt = rFontIds.begin(); aIt->mnPos < nStrLen; ++aIt )
        {
            sal_Int32 nPortionLen = (aIt + 1)->mnPos - aIt->mnPos;
            if( nPortionLen > 0 )
            {
                // convert byte string to unicode string, using current font encoding
                FontRef xFont = getStyles().getFont( aIt->mnFontId );
                rtl_TextEncoding eTextEnc = xFont.get() ? xFont->getFontEncoding() : eDefaultTextEnc;
                OUString aUniStr = OStringToOUString( rText.copy( aIt->mnPos, nPortionLen ), eTextEnc );
                // create string portion
                RichStringPortionRef xPortion = createPortion();
                xPortion->setText( aUniStr );
                xPortion->setFontId( aIt->mnFontId );
            }
        }
    }
}

void RichString::createBiffPortions( const OUString& rText, BiffRichStringFontIdVec& rFontIds )
{
    maPortions.clear();
    sal_Int32 nStrLen = rText.getLength();
    if( nStrLen > 0 )
    {
        // add leading and trailing string position to ease the following loop
        if( rFontIds.empty() || (rFontIds.front().mnPos > 0) )
            rFontIds.insert( rFontIds.begin(), BiffRichStringFontId( 0, -1 ) );
        if( rFontIds.back().mnPos < nStrLen )
            rFontIds.push_back( BiffRichStringFontId( nStrLen, -1 ) );

        // create all string portions according to the font id vector
        for( BiffRichStringFontIdVec::const_iterator aIt = rFontIds.begin(); aIt->mnPos < nStrLen; ++aIt )
        {
            sal_Int32 nPortionLen = (aIt + 1)->mnPos - aIt->mnPos;
            if( nPortionLen > 0 )
            {
                RichStringPortionRef xPortion = createPortion();
                xPortion->setText( rText.copy( aIt->mnPos, nPortionLen ) );
                xPortion->setFontId( aIt->mnFontId );
            }
        }
    }
}

void RichString::createBiffPhonetics( const OUString& rText, BiffPhoneticPortionVec& rPhonetics, sal_Int32 nBaseLen )
{
    maPhonetics.clear();
    sal_Int32 nStrLen = rText.getLength();
    if( nStrLen > 0 )
    {
        // no portions - assign phonetic text to entire base text
        if( rPhonetics.empty() )
            rPhonetics.push_back( BiffPhoneticPortion( 0, 0, nBaseLen ) );
        // add trailing string position to ease the following loop
        if( rPhonetics.back().mnPos < nStrLen )
            rPhonetics.push_back( BiffPhoneticPortion( nStrLen, nBaseLen, 0 ) );

        // create all phonetic portions according to the portions vector
        for( BiffPhoneticPortionVec::const_iterator aIt = rPhonetics.begin(); aIt->mnPos < nStrLen; ++aIt )
        {
            sal_Int32 nPortionLen = (aIt + 1)->mnPos - aIt->mnPos;
            if( nPortionLen > 0 )
            {
                RichStringPhoneticRef xPhonetic = createPhonetic();
                xPhonetic->setText( rText.copy( aIt->mnPos, nPortionLen ) );
                xPhonetic->setBaseRange( aIt->mnBasePos, aIt->mnBasePos + aIt->mnBaseLen );
            }
        }
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

