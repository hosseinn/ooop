/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: stylesbuffer.cxx,v $
 *
 *  $Revision: 1.1.2.50 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:21 $
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

#include "oox/xls/stylesbuffer.hxx"
#include <com/sun/star/container/XNameContainer.hpp>
#include <com/sun/star/awt/FontDescriptor.hpp>
#include <com/sun/star/awt/FontFamily.hpp>
#include <com/sun/star/awt/FontWeight.hpp>
#include <com/sun/star/awt/FontSlant.hpp>
#include <com/sun/star/awt/FontUnderline.hpp>
#include <com/sun/star/awt/FontStrikeout.hpp>
#include <com/sun/star/awt/XDevice.hpp>
#include <com/sun/star/awt/XFont2.hpp>
#include <com/sun/star/text/WritingMode2.hpp>
#include <com/sun/star/text/XText.hpp>
#include <com/sun/star/style/XStyleFamiliesSupplier.hpp>
#include <com/sun/star/style/XStyle.hpp>
#include <rtl/tencinfo.h>
#include <rtl/ustrbuf.hxx>
#include "oox/core/attributelist.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/contexthelper.hxx"
#include "oox/xls/themebuffer.hxx"
#include "oox/xls/unitconverter.hxx"

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::container::XNameAccess;
using ::com::sun::star::container::XNameContainer;
using ::com::sun::star::awt::FontDescriptor;
using ::com::sun::star::awt::XDevice;
using ::com::sun::star::awt::XFont2;
using ::com::sun::star::table::BorderLine;
using ::com::sun::star::text::XText;
using ::com::sun::star::style::XStyleFamiliesSupplier;
using ::com::sun::star::style::XStyle;
using ::oox::core::AttributeList;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

OoxColor::OoxColor() :
    mfTint( 0.0 ),
    meType( TYPE_RGB ),
    mnValue( 0 )
{
}

OoxColor::OoxColor( Type eType, sal_Int32 nValue, double fTint ) :
    mfTint( fTint ),
    meType( eType ),
    mnValue( nValue )
{
}

void OoxColor::set( Type eType, sal_Int32 nValue, double fTint )
{
    mfTint = fTint;
    meType = eType;
    mnValue = nValue;
}

void OoxColor::importColor( const AttributeList& rAttribs )
{
    mfTint = rAttribs.getDouble( XML_tint, 0.0 );
    if( rAttribs.hasAttribute( XML_rgb ) )
    {
        meType = TYPE_RGB;
        mnValue = rAttribs.getHex( XML_rgb, API_RGB_TRANSPARENT );
    }
    else if( rAttribs.hasAttribute( XML_theme ) )
    {
        meType = TYPE_THEME;
        mnValue = rAttribs.getInteger( XML_theme, -1 );
    }
    else if( rAttribs.hasAttribute( XML_indexed ) )
    {
        meType = TYPE_PALETTE;
        mnValue = rAttribs.getInteger( XML_indexed, -1 );
    }
}

void OoxColor::importColorId( BiffInputStream& rStrm, bool b16Bit )
{
    mfTint = 0.0;
    meType = TYPE_PALETTE;
    mnValue = b16Bit ? rStrm.readuInt16() : rStrm.readuInt8();
}

void OoxColor::importColorRgb( BiffInputStream& rStrm )
{
    mfTint = 0.0;
    meType = TYPE_RGB;
    sal_uInt8 nR, nG, nB;
    rStrm >> nR >> nG >> nB;
    rStrm.ignore( 1 );
    mnValue = nR;
    mnValue <<= 8;
    mnValue |= nG;
    mnValue <<= 8;
    mnValue |= nB;
}

// ============================================================================

namespace {

/** Standard EGA colors, bright. */
#define PALETTE_EGA_COLORS_LIGHT \
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF
/** Standard EGA colors, dark. */
#define PALETTE_EGA_COLORS_DARK \
            0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080

/** Default color table for BIFF2. */
static const sal_Int32 spnDefColors2[] =
{
/*  0 */    PALETTE_EGA_COLORS_LIGHT
};

/** Default color table for BIFF3/BIFF4. */
static const sal_Int32 spnDefColors3[] =
{
/*  0 */    PALETTE_EGA_COLORS_LIGHT,
/*  8 */    PALETTE_EGA_COLORS_LIGHT,
/* 16 */    PALETTE_EGA_COLORS_DARK
};

/** Default color table for BIFF5. */
static const sal_Int32 spnDefColors5[] =
{
/*  0 */    PALETTE_EGA_COLORS_LIGHT,
/*  8 */    PALETTE_EGA_COLORS_LIGHT,
/* 16 */    PALETTE_EGA_COLORS_DARK,
/* 24 */    0x8080FF, 0x802060, 0xFFFFC0, 0xA0E0E0, 0x600080, 0xFF8080, 0x0080C0, 0xC0C0FF,
/* 32 */    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
/* 40 */    0x00CFFF, 0x69FFFF, 0xE0FFE0, 0xFFFF80, 0xA6CAF0, 0xDD9CB3, 0xB38FEE, 0xE3E3E3,
/* 48 */    0x2A6FF9, 0x3FB8CD, 0x488436, 0x958C41, 0x8E5E42, 0xA0627A, 0x624FAC, 0x969696,
/* 56 */    0x1D2FBE, 0x286676, 0x004500, 0x453E01, 0x6A2813, 0x85396A, 0x4A3285, 0x424242
};

/** Default color table for BIFF8/OOX. */
static const sal_Int32 spnDefColors8[] =
{
/*  0 */    PALETTE_EGA_COLORS_LIGHT,
/*  8 */    PALETTE_EGA_COLORS_LIGHT,
/* 16 */    PALETTE_EGA_COLORS_DARK,
/* 24 */    0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
/* 32 */    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
/* 40 */    0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
/* 48 */    0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
/* 56 */    0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333
};

#undef PALETTE_EGA_COLORS_LIGHT
#undef PALETTE_EGA_COLORS_DARK

} // namespace

// ----------------------------------------------------------------------------

ColorPalette::ColorPalette( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    mnWindowColor( ThemeBuffer::getSystemWindowColor() ),
    mnWinTextColor( ThemeBuffer::getSystemWindowTextColor() )
{
    // default colors
    switch( getFilterType() )
    {
        case FILTER_OOX:
            maColors.insert( maColors.begin(), spnDefColors8, STATIC_TABLE_END( spnDefColors8 ) );
            mnAppendIndex = OOX_COLOR_USEROFFSET;
        break;
        case FILTER_BIFF:
            switch( getBiff() )
            {
                case BIFF2: maColors.insert( maColors.begin(), spnDefColors2, STATIC_TABLE_END( spnDefColors2 ) );  break;
                case BIFF3:
                case BIFF4: maColors.insert( maColors.begin(), spnDefColors3, STATIC_TABLE_END( spnDefColors3 ) );  break;
                case BIFF5: maColors.insert( maColors.begin(), spnDefColors5, STATIC_TABLE_END( spnDefColors5 ) );  break;
                case BIFF8: maColors.insert( maColors.begin(), spnDefColors8, STATIC_TABLE_END( spnDefColors8 ) );  break;
                case BIFF_UNKNOWN: break;
            }
            mnAppendIndex = BIFF_COLOR_USEROFFSET;
        break;
        case FILTER_UNKNOWN: break;
    }
}

bool ColorPalette::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( colors ):
            return  (nElement == XLS_TOKEN( indexedColors ));
        case XLS_TOKEN( indexedColors ):
            return  (nElement == XLS_TOKEN( rgbColor ));
    }
    return false;
}

void ColorPalette::appendColor( sal_Int32 nRGBValue )
{
    if( mnAppendIndex < maColors.size() )
        maColors[ mnAppendIndex ] = nRGBValue;
    else
        maColors.push_back( nRGBValue );
    ++mnAppendIndex;
}

void ColorPalette::importPalette( BiffInputStream& rStrm )
{
    sal_uInt16 nCount;
    rStrm >> nCount;
    OSL_ENSURE( rStrm.getRecLeft() == static_cast< sal_uInt32 >( 4 * nCount ),
        "ColorPalette::importPalette - wrong palette size" );

    // fill palette from OOX_COLOR_USEROFFSET
    mnAppendIndex = OOX_COLOR_USEROFFSET;
    OoxColor aColor;
    for( sal_uInt16 nIndex = 0; rStrm.isValid() && (nIndex < nCount); ++nIndex )
    {
        aColor.importColorRgb( rStrm );
        appendColor( aColor.mnValue );
    }
}

sal_Int32 ColorPalette::getColor( sal_Int32 nIndex ) const
{
    sal_Int32 nColor = API_RGB_TRANSPARENT;
    if( (0 <= nIndex) && (static_cast< size_t >( nIndex ) < maColors.size()) )
    {
        nColor = maColors[ nIndex ];
    }
    else switch( nIndex )
    {
        case OOX_COLOR_WINDOWTEXT3:
        case OOX_COLOR_WINDOWTEXT:
        case OOX_COLOR_CHWINDOWTEXT:    nColor = mnWinTextColor;        break;
        case OOX_COLOR_WINDOWBACK3:
        case OOX_COLOR_WINDOWBACK:
        case OOX_COLOR_CHWINDOWBACK:    nColor = mnWindowColor;         break;
//        case OOX_COLOR_BUTTONBACK:
//        case OOX_COLOR_CHBORDERAUTO:
//        case OOX_COLOR_NOTEBACK:
//        case OOX_COLOR_NOTETEXT:
        case OOX_COLOR_FONTAUTO:        nColor = API_RGB_TRANSPARENT;   break;
        default:
            OSL_ENSURE( false, "ColorPalette::getColor - unknown color index" );
    }
    return nColor;
}

// ============================================================================

OoxFontData::OoxFontData() :
    mnFamily( OOX_FONTFAMILY_NONE ),
    mnCharSet( OOX_FONTCHARSET_ANSI ),
    mfHeight( 0.0 ),
    mnUnderline( XML_none ),
    mnEscapement( XML_baseline ),
    mbBold( false ),
    mbItalic( false ),
    mbStrikeout( false ),
    mbOutline( false ),
    mbShadow( false )
{
}

// ============================================================================

Font::Font( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maOoxData( rGlobalData.getTheme().getDefaultFontData() )
{
}

Font::Font( const GlobalDataHelper& rGlobalData, const OoxFontData& rFontData ) :
    GlobalDataHelper( rGlobalData ),
    maOoxData( rFontData )
{
}

bool Font::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( fonts ):
            return  (nElement == XLS_TOKEN( font ));

        case XLS_TOKEN( font ):
            return  (nElement == XLS_TOKEN( name )) ||
                    (nElement == XLS_TOKEN( charset )) ||
                    (nElement == XLS_TOKEN( family )) ||
                    (nElement == XLS_TOKEN( sz )) ||
                    (nElement == XLS_TOKEN( color )) ||
                    (nElement == XLS_TOKEN( u )) ||
                    (nElement == XLS_TOKEN( vertAlign )) ||
                    (nElement == XLS_TOKEN( b )) ||
                    (nElement == XLS_TOKEN( i )) ||
                    (nElement == XLS_TOKEN( outline )) ||
                    (nElement == XLS_TOKEN( shadow )) ||
                    (nElement == XLS_TOKEN( strike ));

        case XLS_TOKEN( rPr ):
            return  (nElement == XLS_TOKEN( rFont )) ||
                    (nElement == XLS_TOKEN( charset )) ||
                    (nElement == XLS_TOKEN( family )) ||
                    (nElement == XLS_TOKEN( sz )) ||
                    (nElement == XLS_TOKEN( color )) ||
                    (nElement == XLS_TOKEN( u )) ||
                    (nElement == XLS_TOKEN( vertAlign )) ||
                    (nElement == XLS_TOKEN( b )) ||
                    (nElement == XLS_TOKEN( i )) ||
                    (nElement == XLS_TOKEN( outline )) ||
                    (nElement == XLS_TOKEN( shadow )) ||
                    (nElement == XLS_TOKEN( strike )) ||
                    (nElement == XLS_TOKEN( vertAlign ));
    }
    return false;
}

void Font::importAttribs( sal_Int32 nElement, const AttributeList& rAttribs )
{
    const OoxFontData& rDefFontData = getTheme().getDefaultFontData();
    switch( nElement )
    {
        case XLS_TOKEN( name ):
        case XLS_TOKEN( rFont ):
            if( rAttribs.hasAttribute( XML_val ) )
                maOoxData.maName = rAttribs.getString( XML_val );
        break;
        case XLS_TOKEN( family ):       maOoxData.mnFamily = rAttribs.getInteger( XML_val, rDefFontData.mnFamily );     break;
        case XLS_TOKEN( charset ):      maOoxData.mnCharSet = rAttribs.getInteger( XML_val, rDefFontData.mnCharSet );   break;
        case XLS_TOKEN( sz ):           maOoxData.mfHeight = rAttribs.getDouble( XML_val, rDefFontData.mfHeight );      break;
        case XLS_TOKEN( color ):        maOoxData.maColor.importColor( rAttribs );                                      break;
        case XLS_TOKEN( u ):            maOoxData.mnUnderline = rAttribs.getToken( XML_val, XML_single );               break;
        case XLS_TOKEN( vertAlign ):    maOoxData.mnEscapement = rAttribs.getToken( XML_val, XML_baseline );            break;
        case XLS_TOKEN( b ):            maOoxData.mbBold = rAttribs.getBool( XML_val, true );                           break;
        case XLS_TOKEN( i ):            maOoxData.mbItalic = rAttribs.getBool( XML_val, true );                         break;
        case XLS_TOKEN( outline ):      maOoxData.mbOutline = rAttribs.getBool( XML_val, true );                        break;
        case XLS_TOKEN( shadow ):       maOoxData.mbShadow = rAttribs.getBool( XML_val, true );                         break;
        case XLS_TOKEN( strike ):       maOoxData.mbStrikeout = rAttribs.getBool( XML_val, true );                      break;
    }
}

void Font::importFont( BiffInputStream& rStrm )
{
    switch( getBiff() )
    {
        case BIFF2:
            importFontData2( rStrm );
            importFontName2( rStrm );
        break;
        case BIFF3:
        case BIFF4:
            importFontData2( rStrm );
            importFontColor( rStrm );
            importFontName2( rStrm );
        break;
        case BIFF5:
            importFontData2( rStrm );
            importFontColor( rStrm );
            importFontData5( rStrm );
            importFontName2( rStrm );
        break;
        case BIFF8:
            importFontData2( rStrm );
            importFontColor( rStrm );
            importFontData5( rStrm );
            importFontName8( rStrm );
        break;
        case BIFF_UNKNOWN: break;
    }
}

void Font::importFontColor( BiffInputStream& rStrm )
{
    maOoxData.maColor.importColorId( rStrm );
}

rtl_TextEncoding Font::getFontEncoding() const
{
    // #i63105# cells use text encoding from FONT record character set
    // #i67768# BIFF2-BIFF4 FONT records do not contain character set
    // #i71033# do not use maApiData, this function is used before finalizeImport()
    rtl_TextEncoding eFontEnc = RTL_TEXTENCODING_DONTKNOW;
    if( (0 <= maOoxData.mnCharSet) && (maOoxData.mnCharSet <= SAL_MAX_UINT8) )
        eFontEnc = rtl_getTextEncodingFromWindowsCharset( static_cast< sal_uInt8 >( maOoxData.mnCharSet ) );
    return (eFontEnc == RTL_TEXTENCODING_DONTKNOW) ? getTextEncoding() : eFontEnc;
}

void Font::finalizeImport()
{
    namespace cssawt = ::com::sun::star::awt;

    // font name
    maApiData.maDesc.Name = maOoxData.maName;

    // font family
    switch( maOoxData.mnFamily )
    {
        case OOX_FONTFAMILY_NONE:           maApiData.maDesc.Family = cssawt::FontFamily::DONTKNOW;     break;
        case OOX_FONTFAMILY_ROMAN:          maApiData.maDesc.Family = cssawt::FontFamily::ROMAN;        break;
        case OOX_FONTFAMILY_SWISS:          maApiData.maDesc.Family = cssawt::FontFamily::SWISS;        break;
        case OOX_FONTFAMILY_MODERN:         maApiData.maDesc.Family = cssawt::FontFamily::MODERN;       break;
        case OOX_FONTFAMILY_SCRIPT:         maApiData.maDesc.Family = cssawt::FontFamily::SCRIPT;       break;
        case OOX_FONTFAMILY_DECORATIVE:     maApiData.maDesc.Family = cssawt::FontFamily::DECORATIVE;   break;
    }

    // character set
    if( (0 <= maOoxData.mnCharSet) && (maOoxData.mnCharSet <= 255) )
        maApiData.maDesc.CharSet = static_cast< sal_Int16 >(
            rtl_getTextEncodingFromWindowsCharset( static_cast< sal_uInt8 >( maOoxData.mnCharSet ) ) );

    // color, height, weight, slant, strikeout, outline, shadow
    maApiData.mnColor = getStyles().getColor( maOoxData.maColor );
    maApiData.maDesc.Height = static_cast< sal_Int16 >( maOoxData.mfHeight * 20.0 );
    maApiData.maDesc.Weight = maOoxData.mbBold ? cssawt::FontWeight::BOLD : cssawt::FontWeight::NORMAL;
    maApiData.maDesc.Slant = maOoxData.mbItalic ? cssawt::FontSlant_ITALIC : cssawt::FontSlant_NONE;
    maApiData.maDesc.Strikeout = maOoxData.mbStrikeout ? cssawt::FontStrikeout::SINGLE : cssawt::FontStrikeout::NONE;
    maApiData.mbOutline = maOoxData.mbOutline;
    maApiData.mbShadow = maOoxData.mbShadow;

    // underline
    switch( maOoxData.mnUnderline )
    {
        case XML_double:            maApiData.maDesc.Underline = cssawt::FontUnderline::DOUBLE; break;
        case XML_doubleAccounting:  maApiData.maDesc.Underline = cssawt::FontUnderline::DOUBLE; break;
        case XML_none:              maApiData.maDesc.Underline = cssawt::FontUnderline::NONE;   break;
        case XML_single:            maApiData.maDesc.Underline = cssawt::FontUnderline::SINGLE; break;
        case XML_singleAccounting:  maApiData.maDesc.Underline = cssawt::FontUnderline::SINGLE; break;
    }

    // escapement
    switch( maOoxData.mnEscapement )
    {
        case XML_baseline:
            maApiData.mnEscapement = API_ESCAPE_NONE;
            maApiData.mnEscapeHeight = API_ESCAPEHEIGHT_NONE;
        break;
        case XML_superscript:
            maApiData.mnEscapement = API_ESCAPE_SUPERSCRIPT;
            maApiData.mnEscapeHeight = API_ESCAPEHEIGHT_DEFAULT;
        break;
        case XML_subscript:
            maApiData.mnEscapement = API_ESCAPE_SUBSCRIPT;
            maApiData.mnEscapeHeight = API_ESCAPEHEIGHT_DEFAULT;
        break;
    }

    // supported script types
    Reference< XDevice > xDevice = getReferenceDevice();
    if( xDevice.is() )
    {
        Reference< XFont2 > xFont( xDevice->getFont( maApiData.maDesc ), UNO_QUERY );
        if( xFont.is() )
        {
            // #91658# CJK fonts
            maApiData.mbHasAsian =
                xFont->hasGlyphs( OUString( sal_Unicode( 0x3041 ) ) ) ||    // 3040-309F: Hiragana
                xFont->hasGlyphs( OUString( sal_Unicode( 0x30A1 ) ) ) ||    // 30A0-30FF: Katakana
                xFont->hasGlyphs( OUString( sal_Unicode( 0x3111 ) ) ) ||    // 3100-312F: Bopomofo
                xFont->hasGlyphs( OUString( sal_Unicode( 0x3131 ) ) ) ||    // 3130-318F: Hangul Compatibility Jamo
                xFont->hasGlyphs( OUString( sal_Unicode( 0x3301 ) ) ) ||    // 3300-33FF: CJK Compatibility
                xFont->hasGlyphs( OUString( sal_Unicode( 0x3401 ) ) ) ||    // 3400-4DBF: CJK Unified Ideographs Extension A
                xFont->hasGlyphs( OUString( sal_Unicode( 0x4E01 ) ) ) ||    // 4E00-9FAF: CJK Unified Ideographs
                xFont->hasGlyphs( OUString( sal_Unicode( 0x7E01 ) ) ) ||    // 4E00-9FAF: CJK unified ideographs
                xFont->hasGlyphs( OUString( sal_Unicode( 0xA001 ) ) ) ||    // A001-A48F: Yi Syllables
                xFont->hasGlyphs( OUString( sal_Unicode( 0xAC01 ) ) ) ||    // AC00-D7AF: Hangul Syllables
                xFont->hasGlyphs( OUString( sal_Unicode( 0xCC01 ) ) ) ||    // AC00-D7AF: Hangul Syllables
                xFont->hasGlyphs( OUString( sal_Unicode( 0xF901 ) ) ) ||    // F900-FAFF: CJK Compatibility Ideographs
                xFont->hasGlyphs( OUString( sal_Unicode( 0xFF71 ) ) );      // FF00-FFEF: Halfwidth/Fullwidth Forms
            // #113783# CTL fonts
            maApiData.mbHasCmplx =
                xFont->hasGlyphs( OUString( sal_Unicode( 0x05D1 ) ) ) ||    // 0590-05FF: Hebrew
                xFont->hasGlyphs( OUString( sal_Unicode( 0x0631 ) ) ) ||    // 0600-06FF: Arabic
                xFont->hasGlyphs( OUString( sal_Unicode( 0x0721 ) ) ) ||    // 0700-074F: Syriac
                xFont->hasGlyphs( OUString( sal_Unicode( 0x0911 ) ) ) ||    // 0900-0DFF: Indic scripts
                xFont->hasGlyphs( OUString( sal_Unicode( 0x0E01 ) ) ) ||    // 0E00-0E7F: Thai
                xFont->hasGlyphs( OUString( sal_Unicode( 0xFB21 ) ) ) ||    // FB1D-FB4F: Hebrew Presentation Forms
                xFont->hasGlyphs( OUString( sal_Unicode( 0xFB51 ) ) ) ||    // FB50-FDFF: Arabic Presentation Forms-A
                xFont->hasGlyphs( OUString( sal_Unicode( 0xFE71 ) ) );      // FE70-FEFF: Arabic Presentation Forms-B
            // Western fonts
            maApiData.mbHasWstrn =
                (!maApiData.mbHasAsian && !maApiData.mbHasCmplx) ||
                xFont->hasGlyphs( OUString( sal_Unicode( 'A' ) ) );
        }
    }
}

const FontDescriptor& Font::getFontDescriptor() const
{
    return maApiData.maDesc;
}

bool Font::needsRichTextFormat() const
{
    return maApiData.mnEscapement != API_ESCAPE_NONE;
}

void Font::writeToPropertySet( PropertySet& rPropSet, FontPropertyType ePropType ) const
{
    getStylesPropertyHelper().writeFontProperties( rPropSet, maApiData, ePropType );
}

void Font::importFontData2( BiffInputStream& rStrm )
{
    sal_uInt16 nHeight, nFlags;
    rStrm >> nHeight >> nFlags;

    maOoxData.mnFamily     = OOX_FONTFAMILY_NONE;
    maOoxData.mnCharSet    = OOX_FONTCHARSET_UNUSED; // ensure to not use font charset in byte string import
    maOoxData.mfHeight     = nHeight / 20.0;
    maOoxData.mnUnderline  = getFlagValue( nFlags, BIFF_FONTFLAG_UNDERLINE, XML_single, XML_none );
    maOoxData.mnEscapement = XML_none;
    maOoxData.mbBold       = getFlag( nFlags, BIFF_FONTFLAG_BOLD );
    maOoxData.mbItalic     = getFlag( nFlags, BIFF_FONTFLAG_ITALIC );
    maOoxData.mbStrikeout  = getFlag( nFlags, BIFF_FONTFLAG_STRIKEOUT );
    maOoxData.mbOutline    = getFlag( nFlags, BIFF_FONTFLAG_OUTLINE );
    maOoxData.mbShadow     = getFlag( nFlags, BIFF_FONTFLAG_SHADOW );
}

void Font::importFontData5( BiffInputStream& rStrm )
{
    sal_uInt16 nWeight, nEscapement;
    sal_uInt8 nUnderline, nFamily, nCharSet;
    rStrm >> nWeight >> nEscapement >> nUnderline >> nFamily >> nCharSet;
    rStrm.ignore( 1 );

    // equal constants in XML and BIFF for family and charset
    maOoxData.mnFamily     = static_cast< sal_Int32 >( nFamily );
    maOoxData.mnCharSet    = static_cast< sal_Int32 >( nCharSet );

    // convert font underline to XML attribute constants
    switch( nUnderline )
    {
        case BIFF_FONTUNDERL_NONE:          maOoxData.mnUnderline = XML_none;               break;
        case BIFF_FONTUNDERL_SINGLE:        maOoxData.mnUnderline = XML_single;             break;
        case BIFF_FONTUNDERL_DOUBLE:        maOoxData.mnUnderline = XML_double;             break;
        case BIFF_FONTUNDERL_SINGLE_ACC:    maOoxData.mnUnderline = XML_singleAccounting;   break;
        case BIFF_FONTUNDERL_DOUBLE_ACC:    maOoxData.mnUnderline = XML_doubleAccounting;   break;
        default:                            maOoxData.mnUnderline = XML_none;
    }

    // convert font escapement to XML attribute constants
    switch( nEscapement )
    {
        case BIFF_FONTESC_NONE:     maOoxData.mnEscapement = XML_baseline;      break;
        case BIFF_FONTESC_SUPER:    maOoxData.mnEscapement = XML_superscript;   break;
        case BIFF_FONTESC_SUB:      maOoxData.mnEscapement = XML_subscript;     break;
        default:                    maOoxData.mnEscapement = XML_baseline;
    }

    // font weight stored as value, not as flag anymore
    maOoxData.mbBold = nWeight > 450;
}

void Font::importFontName2( BiffInputStream& rStrm )
{
    maOoxData.maName = rStrm.readByteString( false, getTextEncoding() );
}

void Font::importFontName8( BiffInputStream& rStrm )
{
    maOoxData.maName = rStrm.readUniString( rStrm.readuInt8() );
}

// ============================================================================

OoxAlignmentData::OoxAlignmentData() :
    mnHorAlign( XML_general ),
    mnVerAlign( XML_bottom ),
    mnTextDir( OOX_XF_TEXTDIR_CONTEXT ),
    mnRotation( OOX_XF_ROTATION_NONE ),
    mnIndent( OOX_XF_INDENT_NONE ),
    mbWrapText( false ),
    mbShrink( false )
{
}

void OoxAlignmentData::setBiffHorAlign( sal_uInt8 nHorAlign )
{
    static const sal_Int32 spnHorAlign[] = {
        XML_general, XML_left, XML_center, XML_right,
        XML_fill, XML_justify, XML_centerContinuous, XML_distributed };
    mnHorAlign = (nHorAlign < STATIC_TABLE_SIZE( spnHorAlign )) ? spnHorAlign[ nHorAlign ] : XML_general;
}

void OoxAlignmentData::setBiffVerAlign( sal_uInt8 nVerAlign )
{
    static const sal_Int32 spnVerAlign[] = {
        XML_top, XML_center, XML_bottom, XML_justify, XML_distributed };
    mnVerAlign = (nVerAlign < STATIC_TABLE_SIZE( spnVerAlign )) ? spnVerAlign[ nVerAlign ] : XML_bottom;
}

void OoxAlignmentData::setBiffTextOrient( sal_uInt8 nTextOrient )
{
    static const sal_Int32 spnRotation[] = {
        OOX_XF_ROTATION_NONE, OOX_XF_ROTATION_STACKED,
        OOX_XF_ROTATION_90CCW, OOX_XF_ROTATION_90CW };
    mnRotation = (nTextOrient < STATIC_TABLE_SIZE( spnRotation )) ? spnRotation[ nTextOrient ] : OOX_XF_ROTATION_NONE;
}

// ============================================================================

Alignment::Alignment( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void Alignment::importAlignment( const AttributeList& rAttribs )
{
    maOoxData.mnHorAlign = rAttribs.getToken( XML_horizontal, XML_general );
    maOoxData.mnVerAlign = rAttribs.getToken( XML_vertical, XML_bottom );
    maOoxData.mnTextDir = rAttribs.getInteger( XML_readingOrder, OOX_XF_TEXTDIR_CONTEXT );
    maOoxData.mnRotation = rAttribs.getInteger( XML_textRotation, OOX_XF_ROTATION_NONE );
    maOoxData.mnIndent = rAttribs.getInteger( XML_indent, OOX_XF_INDENT_NONE );
    maOoxData.mbWrapText = rAttribs.getBool( XML_wrapText, false );
    maOoxData.mbShrink = rAttribs.getBool( XML_shrinkToFit, false );
}

void Alignment::setBiff2Data( sal_uInt8 nFlags )
{
    maOoxData.setBiffHorAlign( extractValue< sal_uInt8 >( nFlags, 0, 3 ) );
}

void Alignment::setBiff3Data( sal_uInt16 nAlign )
{
    maOoxData.setBiffHorAlign( extractValue< sal_uInt8 >( nAlign, 0, 3 ) );
    maOoxData.mbWrapText = getFlag( nAlign, BIFF_XF_LINEBREAK ); // new in BIFF3
}

void Alignment::setBiff4Data( sal_uInt16 nAlign )
{
    maOoxData.setBiffHorAlign( extractValue< sal_uInt8 >( nAlign, 0, 3 ) );
    maOoxData.setBiffVerAlign( extractValue< sal_uInt8 >( nAlign, 4, 2 ) ); // new in BIFF4
    maOoxData.setBiffTextOrient( extractValue< sal_uInt8 >( nAlign, 6, 2 ) ); // new in BIFF4
    maOoxData.mbWrapText = getFlag( nAlign, BIFF_XF_LINEBREAK );
}

void Alignment::setBiff5Data( sal_uInt16 nAlign )
{
    maOoxData.setBiffHorAlign( extractValue< sal_uInt8 >( nAlign, 0, 3 ) );
    maOoxData.setBiffVerAlign( extractValue< sal_uInt8 >( nAlign, 4, 3 ) );
    maOoxData.setBiffTextOrient( extractValue< sal_uInt8 >( nAlign, 8, 2 ) );
    maOoxData.mbWrapText = getFlag( nAlign, BIFF_XF_LINEBREAK );
}

void Alignment::setBiff8Data( sal_uInt16 nAlign, sal_uInt16 nMiscAttrib )
{
    maOoxData.setBiffHorAlign( extractValue< sal_uInt8 >( nAlign, 0, 3 ) );
    maOoxData.setBiffVerAlign( extractValue< sal_uInt8 >( nAlign, 4, 3 ) );
    maOoxData.mnTextDir = extractValue< sal_Int32 >( nMiscAttrib, 6, 2 ); // new in BIFF8
    maOoxData.mnRotation = extractValue< sal_Int32 >( nAlign, 8, 8 ); // new in BIFF8
    maOoxData.mnIndent = extractValue< sal_uInt8 >( nMiscAttrib, 0, 4 ); // new in BIFF8
    maOoxData.mbWrapText = getFlag( nAlign, BIFF_XF_LINEBREAK );
    maOoxData.mbShrink = getFlag( nMiscAttrib, BIFF_XF_SHRINK ); // new in BIFF8
}

void Alignment::finalizeImport()
{
    namespace csstab = ::com::sun::star::table;
    namespace csstxt = ::com::sun::star::text;

    // horizontal alignment
    switch( maOoxData.mnHorAlign )
    {
        case XML_center:            maApiData.meHorJustify = csstab::CellHoriJustify_CENTER;    break;
        case XML_centerContinuous:  maApiData.meHorJustify = csstab::CellHoriJustify_CENTER;    break;
        case XML_distributed:       maApiData.meHorJustify = csstab::CellHoriJustify_BLOCK;     break;
        case XML_fill:              maApiData.meHorJustify = csstab::CellHoriJustify_REPEAT;    break;
        case XML_general:           maApiData.meHorJustify = csstab::CellHoriJustify_STANDARD;  break;
        case XML_justify:           maApiData.meHorJustify = csstab::CellHoriJustify_BLOCK;     break;
        case XML_left:              maApiData.meHorJustify = csstab::CellHoriJustify_LEFT;      break;
        case XML_right:             maApiData.meHorJustify = csstab::CellHoriJustify_RIGHT;     break;
    }

    // vertical alignment
    switch( maOoxData.mnVerAlign )
    {
        case XML_bottom:        maApiData.meVerJustify = csstab::CellVertJustify_BOTTOM;    break;
        case XML_center:        maApiData.meVerJustify = csstab::CellVertJustify_CENTER;    break;
        case XML_distributed:   maApiData.meVerJustify = csstab::CellVertJustify_TOP;       break;
        case XML_justify:       maApiData.meVerJustify = csstab::CellVertJustify_TOP;       break;
        case XML_top:           maApiData.meVerJustify = csstab::CellVertJustify_TOP;       break;
    }

    /*  indentation: expressed as number of blocks of 3 space characters in
        OOX, and as 10 times the number of points in BIFF. */
    sal_Int32 nIndent = 0;
    switch( getFilterType() )
    {
        case FILTER_OOX:    nIndent = getUnitConverter().calcMm100FromSpaces( 3.0 * maOoxData.mnIndent );   break;
        case FILTER_BIFF:   nIndent = getUnitConverter().calcMm100FromPoints( 10.0 * maOoxData.mnIndent );  break;
        case FILTER_UNKNOWN: break;
    }
    if( (0 <= nIndent) && (nIndent <= SAL_MAX_INT16) )
        maApiData.mnIndent = static_cast< sal_Int16 >( nIndent );

    // complex text direction
    switch( maOoxData.mnTextDir )
    {
        case OOX_XF_TEXTDIR_CONTEXT:    maApiData.mnWritingMode = csstxt::WritingMode2::PAGE;   break;
        case OOX_XF_TEXTDIR_LTR:        maApiData.mnWritingMode = csstxt::WritingMode2::LR_TB;  break;
        case OOX_XF_TEXTDIR_RTL:        maApiData.mnWritingMode = csstxt::WritingMode2::RL_TB;  break;
    }

    // rotation: 0-90 means 0 to 90 degrees ccw, 91-180 means 1 to 90 degrees cw, 255 means stacked
    sal_Int32 nOoxRot = maOoxData.mnRotation;
    maApiData.mnRotation = ((0 <= nOoxRot) && (nOoxRot <= 90)) ?
        (100 * nOoxRot) :
        (((91 <= nOoxRot) && (nOoxRot <= 180)) ? (100 * (450 - nOoxRot)) : 0);

    // "Orientation" property used for character stacking
    maApiData.meOrientation = (nOoxRot == OOX_XF_ROTATION_STACKED) ?
        csstab::CellOrientation_STACKED : csstab::CellOrientation_STANDARD;

    // alignment flags
    maApiData.mbWrapText = maOoxData.mbWrapText;
    maApiData.mbShrink = maOoxData.mbShrink;

}

// ============================================================================

OoxProtectionData::OoxProtectionData() :
    mbLocked( true ),   // default in Excel and Calc
    mbHidden( false )
{
}

// ============================================================================

Protection::Protection( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void Protection::importProtection( const AttributeList& rAttribs )
{
    maOoxData.mbLocked = rAttribs.getBool( XML_locked, true );
    maOoxData.mbHidden = rAttribs.getBool( XML_hidden, false );
}

void Protection::setBiff2Data( sal_uInt8 nNumFmt )
{
    maOoxData.mbLocked = getFlag( nNumFmt, BIFF2_XF_LOCKED );
    maOoxData.mbHidden = getFlag( nNumFmt, BIFF2_XF_HIDDEN );
}

void Protection::setBiff3Data( sal_uInt16 nProt )
{
    maOoxData.mbLocked = getFlag( nProt, BIFF_XF_LOCKED );
    maOoxData.mbHidden = getFlag( nProt, BIFF_XF_HIDDEN );
}

void Protection::finalizeImport()
{
    maApiData.maCellProt.IsLocked = maOoxData.mbLocked;
    maApiData.maCellProt.IsFormulaHidden = maOoxData.mbHidden;
}

// ============================================================================

OoxBorderLineData::OoxBorderLineData() :
    maColor( OoxColor::TYPE_PALETTE, OOX_COLOR_WINDOWTEXT ),
    mnStyle( XML_none ),
    mbUsed( true )
{
}

void OoxBorderLineData::setBiffData( sal_uInt8 nLineStyle, sal_uInt16 nLineColor )
{
    static const sal_Int32 spnStyleIds[] = {
        XML_none, XML_thin, XML_medium, XML_dashed,
        XML_dotted, XML_thick, XML_double, XML_hair,
        XML_mediumDashed, XML_dashDot, XML_mediumDashDot, XML_dashDotDot,
        XML_mediumDashDotDot, XML_slantDashDot };

    maColor.set( OoxColor::TYPE_PALETTE, nLineColor );
    mnStyle = (nLineStyle < STATIC_TABLE_SIZE( spnStyleIds )) ? spnStyleIds[ nLineStyle ] : XML_none;
    mbUsed = true;
}

// ============================================================================

OoxBorderData::OoxBorderData() :
    mbDiagTLtoBR( false ),
    mbDiagBLtoTR( false )
{
}

// ============================================================================

namespace {

inline void lclSetBorderLineWidth( BorderLine& rBorderLine,
        sal_Int16 nOuter, sal_Int16 nDist = API_LINE_NONE, sal_Int16 nInner = API_LINE_NONE )
{
    rBorderLine.OuterLineWidth = nOuter;
    rBorderLine.LineDistance = nDist;
    rBorderLine.InnerLineWidth = nInner;
}

inline sal_Int32 lclGetBorderLineWidth( const BorderLine& rBorderLine )
{
    return rBorderLine.OuterLineWidth + rBorderLine.LineDistance + rBorderLine.InnerLineWidth;
}

const BorderLine* lclGetThickerLine( const BorderLine& rBorderLine1, sal_Bool bValid1, const BorderLine& rBorderLine2, sal_Bool bValid2 )
{
    if( bValid1 && bValid2 )
        return (lclGetBorderLineWidth( rBorderLine1 ) < lclGetBorderLineWidth( rBorderLine2 )) ? &rBorderLine2 : &rBorderLine1;
    if( bValid1 )
        return &rBorderLine1;
    if( bValid2 )
        return &rBorderLine2;
    return 0;
}

} // namespace

// ----------------------------------------------------------------------------

Border::Border( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

bool Border::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( borders ):
            return  (nElement == XLS_TOKEN( border ));
        case XLS_TOKEN( border ):
            return  (nElement == XLS_TOKEN( left )) ||
                    (nElement == XLS_TOKEN( right )) ||
                    (nElement == XLS_TOKEN( top )) ||
                    (nElement == XLS_TOKEN( bottom )) ||
                    (nElement == XLS_TOKEN( diagonal ));
        case XLS_TOKEN( left ):
        case XLS_TOKEN( right ):
        case XLS_TOKEN( top ):
        case XLS_TOKEN( bottom ):
        case XLS_TOKEN( diagonal ):
            return  (nElement == XLS_TOKEN( color ));
    }
    return false;
}

void Border::importBorder( const AttributeList& rAttribs )
{
    maOoxData.mbDiagTLtoBR = rAttribs.getBool( XML_diagonalDown, false );
    maOoxData.mbDiagBLtoTR = rAttribs.getBool( XML_diagonalUp, false );
}

void Border::importStyle( sal_Int32 nElement, const AttributeList& rAttribs )
{
    if( OoxBorderLineData* pBorderLine = getBorderLine( nElement ) )
        pBorderLine->mnStyle = rAttribs.getToken( XML_style, XML_none );
}

void Border::importColor( sal_Int32 nElement, const AttributeList& rAttribs )
{
    if( OoxBorderLineData* pBorderLine = getBorderLine( nElement ) )
        pBorderLine->maColor.importColor( rAttribs );
}

void Border::setBiff2Data( sal_uInt8 nFlags )
{
    maOoxData.maLeft.setBiffData(   getFlagValue( nFlags, BIFF2_XF_LEFTLINE,   BIFF_LINE_THIN, BIFF_LINE_NONE ), BIFF2_COLOR_BLACK );
    maOoxData.maRight.setBiffData(  getFlagValue( nFlags, BIFF2_XF_RIGHTLINE,  BIFF_LINE_THIN, BIFF_LINE_NONE ), BIFF2_COLOR_BLACK );
    maOoxData.maTop.setBiffData(    getFlagValue( nFlags, BIFF2_XF_TOPLINE,    BIFF_LINE_THIN, BIFF_LINE_NONE ), BIFF2_COLOR_BLACK );
    maOoxData.maBottom.setBiffData( getFlagValue( nFlags, BIFF2_XF_BOTTOMLINE, BIFF_LINE_THIN, BIFF_LINE_NONE ), BIFF2_COLOR_BLACK );
    maOoxData.maDiagonal.mbUsed = false;
}

void Border::setBiff3Data( sal_uInt32 nBorder )
{
    maOoxData.maLeft.setBiffData(   extractValue< sal_uInt8 >( nBorder,  8, 3 ), extractValue< sal_uInt16 >( nBorder, 11, 5 ) );
    maOoxData.maRight.setBiffData(  extractValue< sal_uInt8 >( nBorder, 24, 3 ), extractValue< sal_uInt16 >( nBorder, 27, 5 ) );
    maOoxData.maTop.setBiffData(    extractValue< sal_uInt8 >( nBorder,  0, 3 ), extractValue< sal_uInt16 >( nBorder,  3, 5 ) );
    maOoxData.maBottom.setBiffData( extractValue< sal_uInt8 >( nBorder, 16, 3 ), extractValue< sal_uInt16 >( nBorder, 19, 5 ) );
    maOoxData.maDiagonal.mbUsed = false;
}

void Border::setBiff5Data( sal_uInt32 nBorder, sal_uInt32 nArea )
{
    maOoxData.maLeft.setBiffData(   extractValue< sal_uInt8 >( nBorder,  3, 3 ), extractValue< sal_uInt16 >( nBorder, 16, 7 ) );
    maOoxData.maRight.setBiffData(  extractValue< sal_uInt8 >( nBorder,  6, 3 ), extractValue< sal_uInt16 >( nBorder, 23, 7 ) );
    maOoxData.maTop.setBiffData(    extractValue< sal_uInt8 >( nBorder,  0, 3 ), extractValue< sal_uInt16 >( nBorder,  9, 7 ) );
    maOoxData.maBottom.setBiffData( extractValue< sal_uInt8 >( nArea,   22, 3 ), extractValue< sal_uInt16 >( nArea,   25, 7 ) );
    maOoxData.maDiagonal.mbUsed = false;
}

void Border::setBiff8Data( sal_uInt32 nBorder1, sal_uInt32 nBorder2 )
{
    maOoxData.maLeft.setBiffData(   extractValue< sal_uInt8 >( nBorder1,  0, 4 ), extractValue< sal_uInt16 >( nBorder1, 16, 7 ) );
    maOoxData.maRight.setBiffData(  extractValue< sal_uInt8 >( nBorder1,  4, 4 ), extractValue< sal_uInt16 >( nBorder1, 23, 7 ) );
    maOoxData.maTop.setBiffData(    extractValue< sal_uInt8 >( nBorder1,  8, 4 ), extractValue< sal_uInt16 >( nBorder2,  0, 7 ) );
    maOoxData.maBottom.setBiffData( extractValue< sal_uInt8 >( nBorder1, 12, 4 ), extractValue< sal_uInt16 >( nBorder2,  7, 7 ) );
    maOoxData.mbDiagTLtoBR = getFlag( nBorder1, BIFF_XF_DIAG_TLBR );
    maOoxData.mbDiagBLtoTR = getFlag( nBorder1, BIFF_XF_DIAG_BLTR );
    if( maOoxData.mbDiagTLtoBR || maOoxData.mbDiagBLtoTR )
        maOoxData.maDiagonal.setBiffData( extractValue< sal_uInt8 >( nBorder2, 21, 4 ), extractValue< sal_uInt16 >( nBorder2, 14, 7 ) );
}

void Border::finalizeImport()
{
    maApiData.mbLeftUsed   = maOoxData.maLeft.mbUsed;
    maApiData.mbRightUsed  = maOoxData.maRight.mbUsed;
    maApiData.mbTopUsed    = maOoxData.maTop.mbUsed;
    maApiData.mbBottomUsed = maOoxData.maBottom.mbUsed;
    maApiData.mbDiagUsed   = maOoxData.maDiagonal.mbUsed;

    maApiData.maBorder.IsLeftLineValid   = convertBorderLine( maApiData.maBorder.LeftLine,   maOoxData.maLeft );
    maApiData.maBorder.IsRightLineValid  = convertBorderLine( maApiData.maBorder.RightLine,  maOoxData.maRight );
    maApiData.maBorder.IsTopLineValid    = convertBorderLine( maApiData.maBorder.TopLine,    maOoxData.maTop );
    maApiData.maBorder.IsBottomLineValid = convertBorderLine( maApiData.maBorder.BottomLine, maOoxData.maBottom );

    maApiData.maBorder.IsVerticalLineValid = maApiData.maBorder.IsLeftLineValid || maApiData.maBorder.IsRightLineValid;
    if( const BorderLine* pVertLine = lclGetThickerLine( maApiData.maBorder.LeftLine, maApiData.maBorder.IsLeftLineValid, maApiData.maBorder.RightLine, maApiData.maBorder.IsRightLineValid ) )
        maApiData.maBorder.VerticalLine = *pVertLine;

    maApiData.maBorder.IsHorizontalLineValid = maApiData.maBorder.IsTopLineValid || maApiData.maBorder.IsBottomLineValid;
    if( const BorderLine* pHorLine = lclGetThickerLine( maApiData.maBorder.TopLine, maApiData.maBorder.IsTopLineValid, maApiData.maBorder.BottomLine, maApiData.maBorder.IsBottomLineValid ) )
        maApiData.maBorder.HorizontalLine = *pHorLine;

    if( maOoxData.mbDiagTLtoBR )
        convertBorderLine( maApiData.maTLtoBR, maOoxData.maDiagonal );
    if( maOoxData.mbDiagBLtoTR )
        convertBorderLine( maApiData.maBLtoTR, maOoxData.maDiagonal );
}

void Border::writeToPropertySet( PropertySet& rPropSet ) const
{
    getStylesPropertyHelper().writeBorderProperties( rPropSet, maApiData );
}

OoxBorderLineData* Border::getBorderLine( sal_Int32 nElement )
{
    switch( nElement )
    {
        case XLS_TOKEN( left ):     return &maOoxData.maLeft;
        case XLS_TOKEN( right ):    return &maOoxData.maRight;
        case XLS_TOKEN( top ):      return &maOoxData.maTop;
        case XLS_TOKEN( bottom ):   return &maOoxData.maBottom;
        case XLS_TOKEN( diagonal ): return &maOoxData.maDiagonal;
    }
    return 0;
}

bool Border::convertBorderLine( BorderLine& rBorderLine, const OoxBorderLineData& rLineData )
{
    rBorderLine.Color = getStyles().getColor( rLineData.maColor );
    switch( rLineData.mnStyle )
    {
        case XML_dashDot:           lclSetBorderLineWidth( rBorderLine, API_LINE_THIN );    break;
        case XML_dashDotDot:        lclSetBorderLineWidth( rBorderLine, API_LINE_THIN );    break;
        case XML_dashed:            lclSetBorderLineWidth( rBorderLine, API_LINE_THIN );    break;
        case XML_dotted:            lclSetBorderLineWidth( rBorderLine, API_LINE_THIN );    break;
        case XML_double:            lclSetBorderLineWidth( rBorderLine, API_LINE_THIN, API_LINE_THIN, API_LINE_THIN ); break;
        case XML_hair:              lclSetBorderLineWidth( rBorderLine, API_LINE_HAIR );    break;
        case XML_medium:            lclSetBorderLineWidth( rBorderLine, API_LINE_MEDIUM );  break;
        case XML_mediumDashDot:     lclSetBorderLineWidth( rBorderLine, API_LINE_MEDIUM );  break;
        case XML_mediumDashDotDot:  lclSetBorderLineWidth( rBorderLine, API_LINE_MEDIUM );  break;
        case XML_mediumDashed:      lclSetBorderLineWidth( rBorderLine, API_LINE_MEDIUM );  break;
        case XML_none:              lclSetBorderLineWidth( rBorderLine, API_LINE_NONE );    break;
        case XML_slantDashDot:      lclSetBorderLineWidth( rBorderLine, API_LINE_MEDIUM );  break;
        case XML_thick:             lclSetBorderLineWidth( rBorderLine, API_LINE_THICK );   break;
        case XML_thin:              lclSetBorderLineWidth( rBorderLine, API_LINE_THIN );    break;
        default:                    lclSetBorderLineWidth( rBorderLine, API_LINE_NONE );    break;
    }
    return rLineData.mbUsed;
}


// ============================================================================

OoxPatternFillData::OoxPatternFillData( bool bUsedFlagsDefault ) :
    maPatternColor( OoxColor::TYPE_PALETTE, OOX_COLOR_WINDOWTEXT ),
    maFillColor( OoxColor::TYPE_PALETTE, OOX_COLOR_WINDOWBACK ),
    mnPattern( XML_none ),
    mbPatternUsed( bUsedFlagsDefault ),
    mbPattColorUsed( bUsedFlagsDefault ),
    mbFillColorUsed( bUsedFlagsDefault )
{
}

void OoxPatternFillData::setBiffData( sal_uInt16 nPatternColor, sal_uInt16 nFillColor, sal_uInt8 nPattern )
{
    static const sal_Int32 spnPatternIds[] = {
        XML_none, XML_solid, XML_mediumGray, XML_darkGray,
        XML_lightGray, XML_darkHorizontal, XML_darkVertical, XML_darkDown,
        XML_darkUp, XML_darkGrid, XML_darkTrellis, XML_lightHorizontal,
        XML_lightVertical, XML_lightDown, XML_lightUp, XML_lightGrid,
        XML_lightTrellis, XML_gray125, XML_gray0625 };

    maPatternColor.set( OoxColor::TYPE_PALETTE, static_cast< sal_Int32 >( nPatternColor ) );
    maFillColor.set( OoxColor::TYPE_PALETTE, static_cast< sal_Int32 >( nFillColor ) );
    mnPattern = (nPattern < STATIC_TABLE_SIZE( spnPatternIds )) ? spnPatternIds[ nPattern ] : XML_none;
}

// ----------------------------------------------------------------------------

OoxGradientFillData::OoxGradientFillData()
{
}

// ============================================================================

namespace {

inline sal_Int32 lclGetMixedColorComp( sal_Int32 nPatt, sal_Int32 nFill, sal_Int32 nAlpha )
{
    return ((nPatt - nFill) * nAlpha) / 0x80 + nFill;
}

sal_Int32 lclGetMixedColor( sal_Int32 nPattColor, sal_Int32 nFillColor, sal_Int32 nAlpha )
{
    return
        (lclGetMixedColorComp( nPattColor & 0xFF0000, nFillColor & 0xFF0000, nAlpha ) & 0xFF0000) |
        (lclGetMixedColorComp( nPattColor & 0x00FF00, nFillColor & 0x00FF00, nAlpha ) & 0x00FF00) |
        (lclGetMixedColorComp( nPattColor & 0x0000FF, nFillColor & 0x0000FF, nAlpha ) & 0x0000FF);
}

} // namespace

// ----------------------------------------------------------------------------

Fill::Fill( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

bool Fill::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( fills ):
            return  (nElement == XLS_TOKEN( fill ));
        case XLS_TOKEN( fill ):
            return  (nElement == XLS_TOKEN( patternFill )) ||
                    (nElement == XLS_TOKEN( gradientFill ));
        case XLS_TOKEN( patternFill ):
            return  (nElement == XLS_TOKEN( fgColor )) ||
                    (nElement == XLS_TOKEN( bgColor ));
        case XLS_TOKEN( gradientFill ):
            return  (nElement == XLS_TOKEN( stop ));
        case XLS_TOKEN( stop ):
            return  (nElement == XLS_TOKEN( color ));
    }
    return false;
}

void Fill::importPatternFill( const AttributeList& rAttribs, bool bUsedFlagsDefault )
{
    mxOoxPattData.reset( new OoxPatternFillData( bUsedFlagsDefault ) );
    mxOoxPattData->mnPattern = rAttribs.getToken( XML_patternType, XML_none );
}

void Fill::importGradientFill( const AttributeList& )
{
    mxOoxGradData.reset( new OoxGradientFillData );
}

void Fill::importFgColor( const AttributeList& rAttribs )
{
    if( mxOoxPattData.get() )
    {
        mxOoxPattData->maPatternColor.importColor( rAttribs );
        mxOoxPattData->mbPattColorUsed = true;
    }
}

void Fill::importBgColor( const AttributeList& rAttribs )
{
    if( mxOoxPattData.get() )
    {
        mxOoxPattData->maFillColor.importColor( rAttribs );
        mxOoxPattData->mbFillColorUsed = true;
    }
}

void Fill::importColor( const AttributeList& rAttribs, sal_Int32 nPosition )
{
    if( mxOoxGradData.get() && (nPosition >= 0) )
    {
        ::boost::shared_ptr< OoxColor > xColor( new OoxColor );
        xColor->importColor( rAttribs );
        mxOoxGradData->maColors[ nPosition ] = xColor;
    }
}

void Fill::setBiff2Data( sal_uInt8 nFlags )
{
    mxOoxPattData.reset( new OoxPatternFillData( true ) );
    mxOoxPattData->setBiffData(
        BIFF2_COLOR_BLACK,
        BIFF2_COLOR_WHITE,
        getFlagValue( nFlags, BIFF2_XF_BACKGROUND, BIFF_PATT_125, BIFF_PATT_NONE ) );
}

void Fill::setBiff3Data( sal_uInt16 nArea )
{
    mxOoxPattData.reset( new OoxPatternFillData( true ) );
    mxOoxPattData->setBiffData(
        extractValue< sal_uInt16 >( nArea, 6, 5 ),
        extractValue< sal_uInt16 >( nArea, 11, 5 ),
        extractValue< sal_uInt8 >( nArea, 0, 6 ) );
}

void Fill::setBiff5Data( sal_uInt32 nArea )
{
    mxOoxPattData.reset( new OoxPatternFillData( true ) );
    mxOoxPattData->setBiffData(
        extractValue< sal_uInt16 >( nArea, 0, 7 ),
        extractValue< sal_uInt16 >( nArea, 7, 7 ),
        extractValue< sal_uInt8 >( nArea, 16, 6 ) );
}

void Fill::setBiff8Data( sal_uInt32 nBorder2, sal_uInt16 nArea )
{
    mxOoxPattData.reset( new OoxPatternFillData( true ) );
    mxOoxPattData->setBiffData(
        extractValue< sal_uInt16 >( nArea, 0, 7 ),
        extractValue< sal_uInt16 >( nArea, 7, 7 ),
        extractValue< sal_uInt8 >( nBorder2, 26, 6 ) );
}

void Fill::finalizeImport()
{
    if( mxOoxPattData.get() )
    {
        // pattern may be unused in cond. formats
//         maApiData.mbUsed = mxOoxPattData->mbPatternUsed;

        if( mxOoxPattData->mnPattern == XML_none )
        {
            if( !mxOoxPattData->mbPatternUsed && mxOoxPattData->mbFillColorUsed )
            {
                // When the pattern type is XML_none and the background (or fill)
                // color exists, then it's a solid fill for conditional formatting.
                maApiData.mnColor = getStyles().getColor( mxOoxPattData->maFillColor );
                maApiData.mbTransparent = false;
            }
            else
            {
                maApiData.mnColor = API_RGB_TRANSPARENT;
                maApiData.mbTransparent = true;
            }
        }
        else
        {
            sal_Int32 nAlpha = 0x80;
            switch( mxOoxPattData->mnPattern )
            {
                case XML_darkDown:          nAlpha = 0x40;  break;
                case XML_darkGray:          nAlpha = 0x60;  break;
                case XML_darkGrid:          nAlpha = 0x40;  break;
                case XML_darkHorizontal:    nAlpha = 0x40;  break;
                case XML_darkTrellis:       nAlpha = 0x60;  break;
                case XML_darkUp:            nAlpha = 0x40;  break;
                case XML_darkVertical:      nAlpha = 0x40;  break;
                case XML_gray0625:          nAlpha = 0x08;  break;
                case XML_gray125:           nAlpha = 0x10;  break;
                case XML_lightDown:         nAlpha = 0x20;  break;
                case XML_lightGray:         nAlpha = 0x20;  break;
                case XML_lightGrid:         nAlpha = 0x38;  break;
                case XML_lightHorizontal:   nAlpha = 0x20;  break;
                case XML_lightTrellis:      nAlpha = 0x30;  break;
                case XML_lightUp:           nAlpha = 0x20;  break;
                case XML_lightVertical:     nAlpha = 0x20;  break;
                case XML_mediumGray:        nAlpha = 0x40;  break;
                case XML_solid:             nAlpha = 0x80;  break;
            }

            sal_Int32 nPattColor = mxOoxPattData->mbPattColorUsed ?
                getStyles().getColor( mxOoxPattData->maPatternColor ) : ThemeBuffer::getSystemWindowTextColor();
            sal_Int32 nFillColor = mxOoxPattData->mbFillColorUsed ?
                getStyles().getColor( mxOoxPattData->maFillColor ) : ThemeBuffer::getSystemWindowColor();

            maApiData.mnColor = lclGetMixedColor( nPattColor, nFillColor, nAlpha );
            maApiData.mbTransparent = false;
        }
    }
    else if( mxOoxGradData.get() && !mxOoxGradData->maColors.empty() )
    {
        maApiData.mnColor = getStyles().getColor( *mxOoxGradData->maColors.begin()->second );
        if( mxOoxGradData->maColors.begin()->second != mxOoxGradData->maColors.rbegin()->second )
        {
            sal_Int32 nEndColor = getStyles().getColor( *mxOoxGradData->maColors.rbegin()->second );
            maApiData.mnColor = lclGetMixedColor( maApiData.mnColor, nEndColor, 0x40 );
            maApiData.mbTransparent = false;
        }
    }
}

void Fill::writeToPropertySet( PropertySet& rPropSet ) const
{
    getStylesPropertyHelper().writeSolidFillProperties( rPropSet, maApiData );
}

// ============================================================================

OoxXfData::OoxXfData() :
    mnStyleXfId( -1 ),
    mnFontId( -1 ),
    mnNumFmtId( -1 ),
    mnBorderId( -1 ),
    mnFillId( -1 ),
    mbCellXf( true ),
    mbAlignUsed( false ),
    mbProtUsed( false ),
    mbFontUsed( false ),
    mbNumFmtUsed( false ),
    mbBorderUsed( false ),
    mbAreaUsed( false )
{
}

// ============================================================================

Xf::Xf( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maAlignment( rGlobalData ),
    maProtection( rGlobalData )
{
}

void Xf::setAllUsedFlags( bool bUsed )
{
    maOoxData.mbAlignUsed = maOoxData.mbProtUsed = maOoxData.mbFontUsed =
        maOoxData.mbNumFmtUsed = maOoxData.mbBorderUsed = maOoxData.mbAreaUsed = bUsed;
}

bool Xf::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( cellStyleXfs ):
        case XLS_TOKEN( cellXfs ):
            return  (nElement == XLS_TOKEN( xf ));
        case XLS_TOKEN( xf ):
            return  (nElement == XLS_TOKEN( alignment )) ||
                    (nElement == XLS_TOKEN( protection ));
    }
    return false;
}

void Xf::importCellXf( const AttributeList& rAttribs, bool bCellXf )
{
    maOoxData.mbCellXf = bCellXf;
    maOoxData.mnStyleXfId = rAttribs.getInteger( XML_xfId, -1 );
    maOoxData.mnFontId = rAttribs.getInteger( XML_fontId, -1 );
    maOoxData.mnNumFmtId = rAttribs.getInteger( XML_numFmtId, -1 );
    maOoxData.mnBorderId = rAttribs.getInteger( XML_borderId, -1 );
    maOoxData.mnFillId = rAttribs.getInteger( XML_fillId, -1 );

    /*  Default value of the apply*** attributes is dependent on context:
        true in cellStyleXfs element, false in cellXfs element... */
    maOoxData.mbAlignUsed  = rAttribs.getBool( XML_applyAlignment,    !maOoxData.mbCellXf );
    maOoxData.mbProtUsed   = rAttribs.getBool( XML_applyProtection,   !maOoxData.mbCellXf );
    maOoxData.mbFontUsed   = rAttribs.getBool( XML_applyFont,         !maOoxData.mbCellXf );
    maOoxData.mbNumFmtUsed = rAttribs.getBool( XML_applyNumberFormat, !maOoxData.mbCellXf );
    maOoxData.mbBorderUsed = rAttribs.getBool( XML_applyBorder,       !maOoxData.mbCellXf );
    maOoxData.mbAreaUsed   = rAttribs.getBool( XML_applyFill,         !maOoxData.mbCellXf );
}

void Xf::importAlignment( const AttributeList& rAttribs )
{
    maAlignment.importAlignment( rAttribs );
}

void Xf::importProtection( const AttributeList& rAttribs )
{
    maProtection.importProtection( rAttribs );
}

void Xf::importXf( BiffInputStream& rStrm )
{
    BorderRef xBorder = getStyles().createBorder( &maOoxData.mnBorderId );
    FillRef xFill = getStyles().createFill( &maOoxData.mnFillId );

    switch( getBiff() )
    {
        case BIFF2:
        {
            sal_uInt8 nFontId, nNumFmtId, nFlags;
            rStrm >> nFontId;
            rStrm.ignore( 1 );
            rStrm >> nNumFmtId >> nFlags;

            // only cell XFs in BIFF2, no parent style, used flags always true
            setAllUsedFlags( true );

            // attributes
            maAlignment.setBiff2Data( nFlags );
            maProtection.setBiff2Data( nNumFmtId );
            xBorder->setBiff2Data( nFlags );
            xFill->setBiff2Data( nFlags );
            maOoxData.mnFontId = static_cast< sal_Int32 >( nFontId );
            maOoxData.mnNumFmtId = static_cast< sal_Int32 >( nNumFmtId & BIFF2_XF_VALFMT_MASK );
        }
        break;

        case BIFF3:
        {
            sal_uInt32 nBorder;
            sal_uInt16 nTypeProt, nAlign, nArea;
            sal_uInt8 nFontId, nNumFmtId;
            rStrm >> nFontId >> nNumFmtId >> nTypeProt >> nAlign >> nArea >> nBorder;

            // XF type/parent
            maOoxData.mbCellXf = !getFlag( nTypeProt, BIFF_XF_STYLE ); // new in BIFF3
            maOoxData.mnStyleXfId = extractValue< sal_Int32 >( nAlign, 4, 12 ); // new in BIFF3
            // attribute used flags
            setUsedFlags( extractValue< sal_uInt8 >( nTypeProt, 10, 6 ) ); // new in BIFF3

            // attributes
            maAlignment.setBiff3Data( nAlign );
            maProtection.setBiff3Data( nTypeProt );
            xBorder->setBiff3Data( nBorder );
            xFill->setBiff3Data( nArea );
            maOoxData.mnFontId = static_cast< sal_Int32 >( nFontId );
            maOoxData.mnNumFmtId = static_cast< sal_Int32 >( nNumFmtId );
        }
        break;

        case BIFF4:
        {
            sal_uInt32 nBorder;
            sal_uInt16 nTypeProt, nAlign, nArea;
            sal_uInt8 nFontId, nNumFmtId;
            rStrm >> nFontId >> nNumFmtId >> nTypeProt >> nAlign >> nArea >> nBorder;

            // XF type/parent
            maOoxData.mbCellXf = !getFlag( nTypeProt, BIFF_XF_STYLE );
            maOoxData.mnStyleXfId = extractValue< sal_Int32 >( nTypeProt, 4, 12 );
            // attribute used flags
            setUsedFlags( extractValue< sal_uInt8 >( nAlign, 10, 6 ) );

            // attributes
            maAlignment.setBiff4Data( nAlign );
            maProtection.setBiff3Data( nTypeProt );
            xBorder->setBiff3Data( nBorder );
            xFill->setBiff3Data( nArea );
            maOoxData.mnFontId = static_cast< sal_Int32 >( nFontId );
            maOoxData.mnNumFmtId = static_cast< sal_Int32 >( nNumFmtId );
        }
        break;

        case BIFF5:
        {
            sal_uInt32 nArea, nBorder;
            sal_uInt16 nFontId, nNumFmtId, nTypeProt, nAlign;
            rStrm >> nFontId >> nNumFmtId >> nTypeProt >> nAlign >> nArea >> nBorder;

            // XF type/parent
            maOoxData.mbCellXf = !getFlag( nTypeProt, BIFF_XF_STYLE );
            maOoxData.mnStyleXfId = extractValue< sal_Int32 >( nTypeProt, 4, 12 );
            // attribute used flags
            setUsedFlags( extractValue< sal_uInt8 >( nAlign, 10, 6 ) );

            // attributes
            maAlignment.setBiff5Data( nAlign );
            maProtection.setBiff3Data( nTypeProt );
            xBorder->setBiff5Data( nBorder, nArea );
            xFill->setBiff5Data( nArea );
            maOoxData.mnFontId = static_cast< sal_Int32 >( nFontId );
            maOoxData.mnNumFmtId = static_cast< sal_Int32 >( nNumFmtId );
        }
        break;

        case BIFF8:
        {
            sal_uInt32 nBorder1, nBorder2;
            sal_uInt16 nFontId, nNumFmtId, nTypeProt, nAlign, nMiscAttrib, nArea;
            rStrm >> nFontId >> nNumFmtId >> nTypeProt >> nAlign >> nMiscAttrib >> nBorder1 >> nBorder2 >> nArea;

            // XF type/parent
            maOoxData.mbCellXf = !getFlag( nTypeProt, BIFF_XF_STYLE );
            maOoxData.mnStyleXfId = extractValue< sal_Int32 >( nTypeProt, 4, 12 );
            // attribute used flags
            setUsedFlags( extractValue< sal_uInt8 >( nMiscAttrib, 10, 6 ) );

            // attributes
            maAlignment.setBiff8Data( nAlign, nMiscAttrib );
            maProtection.setBiff3Data( nTypeProt );
            xBorder->setBiff8Data( nBorder1, nBorder2 );
            xFill->setBiff8Data( nBorder2, nArea );
            maOoxData.mnFontId = static_cast< sal_Int32 >( nFontId );
            maOoxData.mnNumFmtId = static_cast< sal_Int32 >( nNumFmtId );
        }
        break;

        case BIFF_UNKNOWN: break;
    }
}

void Xf::finalizeImport()
{
    // alignment and protection
    maAlignment.finalizeImport();
    maProtection.finalizeImport();
    // update used flags from cell style
    if( maOoxData.mbCellXf )
        if( const Xf* pStyleXf = getStyles().getStyleXf( maOoxData.mnStyleXfId ).get() )
            updateUsedFlags( *pStyleXf );
}

FontRef Xf::getFont() const
{
    return getStyles().getFont( maOoxData.mnFontId );
}

bool Xf::hasAnyUsedFlags() const
{
    return
        maOoxData.mbAlignUsed || maOoxData.mbProtUsed || maOoxData.mbFontUsed ||
        maOoxData.mbNumFmtUsed || maOoxData.mbBorderUsed || maOoxData.mbAreaUsed;
}

void Xf::writeToPropertySet( PropertySet& rPropSet ) const
{
    StylesBuffer& rStyles = getStyles();
    StylesPropertyHelper& rPropHelper = getStylesPropertyHelper();

    // create and set cell style
    if( maOoxData.mbCellXf )
    {
        const OUString& rStyleName = rStyles.createStyleSheet( maOoxData.mnStyleXfId );
        rPropSet.setProperty( CREATE_OUSTRING( "CellStyle" ), rStyleName );
    }

    if( maOoxData.mbFontUsed )
        rStyles.writeFontToPropertySet( rPropSet, maOoxData.mnFontId );
    if( maOoxData.mbNumFmtUsed )
        rStyles.writeNumFmtToPropertySet( rPropSet, maOoxData.mnNumFmtId );
    if( maOoxData.mbAlignUsed )
        rPropHelper.writeAlignmentProperties( rPropSet, maAlignment.getApiData() );
    if( maOoxData.mbProtUsed )
        rPropHelper.writeProtectionProperties( rPropSet, maProtection.getApiData() );
    if( maOoxData.mbBorderUsed )
        rStyles.writeBorderToPropertySet( rPropSet, maOoxData.mnBorderId );
    if( maOoxData.mbAreaUsed )
        rStyles.writeFillToPropertySet( rPropSet, maOoxData.mnFillId );
}

void Xf::setUsedFlags( sal_uInt8 nUsedFlags )
{
    /*  Notes about finding the used flags:
        - In cell XFs a *set* bit means a used attribute.
        - In style XFs a *cleared* bit means a used attribute.
        The boolean flags always store true, if the attribute is used.
        The "maOoxData.mbCellXf == getFlag(...)" construct evaluates to true in
        both mentioned cases: cell XF and set bit; or style XF and cleared bit.
     */
    maOoxData.mbAlignUsed  = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_ALIGN );
    maOoxData.mbProtUsed   = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_PROT );
    maOoxData.mbFontUsed   = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_FONT );
    maOoxData.mbNumFmtUsed = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_VALFMT );
    maOoxData.mbBorderUsed = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_BORDER );
    maOoxData.mbAreaUsed   = maOoxData.mbCellXf == getFlag( nUsedFlags, BIFF_XF_DIFF_AREA );
}

void Xf::updateUsedFlags( const Xf& rStyleXf )
{
    /*  Enables the used flags, if the formatting attributes differ from the
        passed style XF. In cell XFs Excel uses the cell attributes, if they
        differ from the parent style XF.
        #109899# ...or if the respective flag is not set in parent style XF.
     */
    const OoxXfData& rStyleData = rStyleXf.maOoxData;
    if( !maOoxData.mbAlignUsed )
        maOoxData.mbAlignUsed = !rStyleData.mbAlignUsed || !(maAlignment.getApiData() == rStyleXf.maAlignment.getApiData());
    if( !maOoxData.mbProtUsed )
        maOoxData.mbProtUsed = !rStyleData.mbProtUsed || !(maProtection.getApiData() == rStyleXf.maProtection.getApiData());
    if( !maOoxData.mbFontUsed )
        maOoxData.mbFontUsed = !rStyleData.mbFontUsed || (maOoxData.mnFontId != rStyleData.mnFontId);
    if( !maOoxData.mbNumFmtUsed )
        maOoxData.mbNumFmtUsed = !rStyleData.mbNumFmtUsed || (maOoxData.mnNumFmtId != rStyleData.mnNumFmtId);
    if( !maOoxData.mbBorderUsed )
        maOoxData.mbBorderUsed = !rStyleData.mbBorderUsed || (maOoxData.mnBorderId != rStyleData.mnBorderId);
    if( !maOoxData.mbAreaUsed )
        maOoxData.mbAreaUsed = !rStyleData.mbAreaUsed || (maOoxData.mnFillId != rStyleData.mnFillId);
}

// ============================================================================

Dxf::Dxf( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void Dxf::importDxf( const AttributeList& /*rAttribs*/ )
{
    // In fact, dxf elements don't have any attributes according to
    // the spec.
}

FontRef Dxf::getFont()
{
    if ( !mpFont.get() )
        mpFont.reset( new Font( getGlobalData() ) );

    return mpFont;
}

FillRef Dxf::getFill()
{
    if ( !mpFill.get() )
        mpFill.reset( new Fill( getGlobalData() ) );
    return mpFill;
}

BorderRef Dxf::getBorder()
{
    if ( !mpBorder.get() )
        mpBorder.reset( new Border( getGlobalData() ) );
    return mpBorder;
}


// ============================================================================

namespace {

const sal_Char* const spcLegacyStyleNamePrefix = "Excel_BuiltIn_";
const sal_Char* const sppcLegacyStyleNames[] =
{
    "",                     // use existing "Default" style
    "RowLevel_",            // outline level will be appended
    "ColumnLevel_",         // outline level will be appended
    "Comma",
    "Currency",
    "Percent",
    "Comma_0",              // new in BIFF4
    "Currency_0",
    "Hyperlink",            // new in BIFF8
    "Followed_Hyperlink"
};
const sal_Int32 snLegacyStyleNamesCount = static_cast< sal_Int32 >( STATIC_TABLE_SIZE( sppcLegacyStyleNames ) );

const sal_Char* const spcStyleNamePrefix = "Excel Built-in ";
const sal_Char* const sppcStyleNames[] =
{
    "",                     // use existing "Default" style
    "RowLevel_",            // outline level will be appended
    "ColLevel_",            // outline level will be appended
    "Comma",
    "Currency",
    "Percent",
    "Comma [0]",            // new in BIFF4
    "Currency [0]",
    "Hyperlink",            // new in BIFF8
    "Followed Hyperlink",
    "Note",                 // new in OOX
    "Warning Text",
    "",
    "",
    "",
    "Title",
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
    "Input",
    "Output",
    "Calculation",
    "Check Cell",
    "Linked Cell",
    "Total",
    "Good",
    "Bad",
    "Neutral",
    "Accent1",
    "20% - Accent1",
    "40% - Accent1",
    "60% - Accent1",
    "Accent2",
    "20% - Accent2",
    "40% - Accent2",
    "60% - Accent2",
    "Accent3",
    "20% - Accent3",
    "40% - Accent3",
    "60% - Accent3",
    "Accent4",
    "20% - Accent4",
    "40% - Accent4",
    "60% - Accent4",
    "Accent5",
    "20% - Accent5",
    "40% - Accent5",
    "60% - Accent5",
    "Accent6",
    "20% - Accent6",
    "40% - Accent6",
    "60% - Accent6",
    "Explanatory Text"
};
const sal_Int32 snStyleNamesCount = static_cast< sal_Int32 >( STATIC_TABLE_SIZE( sppcStyleNames ) );

const sal_Char* const spcDefaultStyleName = "Default";

OUString lclGetBuiltinStyleName( sal_Int32 nBuiltinId, const OUString& rName, sal_Int32 nLevel = OOX_STYLE_NOLEVEL )
{
    OUStringBuffer aStyleName;
    OSL_ENSURE( (0 <= nBuiltinId) && (nBuiltinId < snStyleNamesCount), "lclGetBuiltinStyleName - unknown builtin style" );
    if( nBuiltinId == OOX_STYLE_NORMAL )    // "Normal" becomes "Default" style
    {
        aStyleName.appendAscii( spcDefaultStyleName );
    }
    else
    {
        aStyleName.appendAscii( spcStyleNamePrefix );
        if( (0 <= nBuiltinId) && (nBuiltinId < snStyleNamesCount) && (sppcStyleNames[ nBuiltinId ][ 0 ] != 0) )
            aStyleName.appendAscii( sppcStyleNames[ nBuiltinId ] );
        else if( rName.getLength() > 0 )
            aStyleName.append( rName );
        else
            aStyleName.append( nBuiltinId );
        if( (nBuiltinId == OOX_STYLE_ROWLEVEL) || (nBuiltinId == OOX_STYLE_COLLEVEL) )
            aStyleName.append( nLevel );
    }
    return aStyleName.makeStringAndClear();
}

bool lclIsBuiltinStyleName( const OUString& rStyleName, sal_Int32* pnBuiltinId, sal_Int32* pnNextChar )
{
    // "Default" becomes "Normal"
    if( rStyleName.equalsIgnoreAsciiCaseAscii( spcDefaultStyleName ) )
    {
        if( pnBuiltinId ) *pnBuiltinId = OOX_STYLE_NORMAL;
        if( pnNextChar ) *pnNextChar = rStyleName.getLength();
        return true;
    }

    // try the other builtin styles
    OUString aPrefix = OUString::createFromAscii( spcStyleNamePrefix );
    sal_Int32 nPrefixLen = aPrefix.getLength();
    sal_Int32 nFoundId = 0;
    sal_Int32 nNextChar = 0;
    if( rStyleName.matchIgnoreAsciiCase( aPrefix ) )
    {
        OUString aShortName;
        for( sal_Int32 nId = 0; nId < snStyleNamesCount; ++nId )
        {
            if( nId != OOX_STYLE_NORMAL )
            {
                aShortName = OUString::createFromAscii( sppcStyleNames[ nId ] );
                if( rStyleName.matchIgnoreAsciiCase( aShortName, nPrefixLen ) &&
                        (nNextChar < nPrefixLen + aShortName.getLength()) )
                {
                    nFoundId = nId;
                    nNextChar = nPrefixLen + aShortName.getLength();
                }
            }
        }
    }

    if( nNextChar > 0 )
    {
        if( pnBuiltinId ) *pnBuiltinId = nFoundId;
        if( pnNextChar ) *pnNextChar = nNextChar;
        return true;
    }

    if( pnBuiltinId ) *pnBuiltinId = OOX_STYLE_USERDEF;
    if( pnNextChar ) *pnNextChar = 0;
    return false;
}

bool lclGetBuiltinStyleId( sal_Int32& rnBuiltinId, sal_Int32& rnLevel, const OUString& rStyleName )
{
    sal_Int32 nBuiltinId;
    sal_Int32 nNextChar;
    if( lclIsBuiltinStyleName( rStyleName, &nBuiltinId, &nNextChar ) )
    {
        if( (nBuiltinId == OOX_STYLE_ROWLEVEL) || (nBuiltinId == OOX_STYLE_COLLEVEL) )
        {
            OUString aLevel = rStyleName.copy( nNextChar );
            sal_Int32 nLevel = aLevel.toInt32();
            if( (0 < nLevel) && (nLevel <= OOX_STYLE_LEVELCOUNT) )
            {
                rnBuiltinId = nBuiltinId;
                rnLevel = nLevel;
                return true;
            }
        }
        else if( rStyleName.getLength() == nNextChar )
        {
            rnBuiltinId = nBuiltinId;
            rnLevel = OOX_STYLE_NOLEVEL;
            return true;
        }
    }
    rnBuiltinId = OOX_STYLE_USERDEF;
    rnLevel = OOX_STYLE_NOLEVEL;
    return false;
}

} // namespace

// ----------------------------------------------------------------------------

OoxCellStyleData::OoxCellStyleData() :
    mnBuiltinId( OOX_STYLE_USERDEF ),
    mnLevel( OOX_STYLE_NOLEVEL ),
    mbCustom( false ),
    mbHidden( false )
{
}

OUString OoxCellStyleData::createStyleName() const
{
    return isBuiltin() ? lclGetBuiltinStyleName( mnBuiltinId, maName, mnLevel ) : maName;
}

// ============================================================================

CellStyle::CellStyle( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

bool CellStyle::isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext )
{
    switch( nParentContext )
    {
        case XLS_TOKEN( cellStyles ):
            return  (nElement == XLS_TOKEN( cellStyle ));
    }
    return false;
}

void CellStyle::importCellStyle( const AttributeList& rAttribs )
{
    maOoxData.maName = rAttribs.getString( XML_name );
    maOoxData.mnBuiltinId = rAttribs.getInteger( XML_builtinId, OOX_STYLE_USERDEF );
    maOoxData.mnLevel = rAttribs.getInteger( XML_iLevel, OOX_STYLE_NOLEVEL );
    maOoxData.mbCustom = rAttribs.getBool( XML_customBuiltin, false );
    maOoxData.mbHidden = rAttribs.getBool( XML_hidden, false );
}

void CellStyle::importStyle( BiffInputStream& rStrm, bool bBuiltin )
{
    if( bBuiltin )
    {
        sal_uInt8 nStyleId, nLevel;
        rStrm >> nStyleId >> nLevel;
        maOoxData.mnBuiltinId = nStyleId;
        maOoxData.mnLevel = nLevel;
    }
    else
    {
        if( getBiff() == BIFF8 )
            maOoxData.maName = rStrm.readUniString();
        else
            maOoxData.maName = rStrm.readByteString( false, getTextEncoding() );
    }
}

bool CellStyle::isDefaultStyle() const
{
    return maOoxData.isDefaultStyle();
}

const OUString& CellStyle::createStyleSheet( sal_Int32 nXfId, bool bSkipDefaultBuiltin )
{
    if( maFinalName.getLength() == 0 )
    {
        bool bBuiltin = maOoxData.isBuiltin();
        if( !bSkipDefaultBuiltin || !bBuiltin || maOoxData.mbCustom )
        {
            // name of the style (generate unique name for builtin styles)
            OUString aStyleName = maOoxData.createStyleName();
            // #i1624# #i1768# ignore unnamed user styles
            if( aStyleName.getLength() > 0 )
            {
                try
                {
                    Reference< XStyleFamiliesSupplier > xFamiliesSup( getDocument(), UNO_QUERY_THROW );
                    Reference< XNameAccess > xFamiliesNA( xFamiliesSup->getStyleFamilies(), UNO_QUERY_THROW );
                    Reference< XNameContainer > xStylesNC( xFamiliesNA->getByName( CREATE_OUSTRING( "CellStyles" ) ), UNO_QUERY_THROW );
                    Reference< XStyle > xStyle;

                    // special handling for default style (do not recreate, but use existing)
                    if( isDefaultStyle() )
                    {
                        /*  Set all flags to true to have all properties in the style,
                            even if the used flags are not set (that's what Excel does). */
                        if( Xf* pXf = getStyles().getStyleXf( nXfId ).get() )
                            pXf->setAllUsedFlags( true );
                        // use existing built-in style
                        xStyle.set( xStylesNC->getByName( aStyleName ), UNO_QUERY_THROW );
                        maFinalName = aStyleName;
                    }
                    else
                    {
                        Reference< XMultiServiceFactory > xFactory( getDocument(), UNO_QUERY_THROW );
                        xStyle.set( xFactory->createInstance( CREATE_OUSTRING( "com.sun.star.style.CellStyle" ) ), UNO_QUERY_THROW );
                        /*  Insert into cell styles collection, passing 'bBuiltin' renames
                            user style, if builtin style with the same name exists. */
                        maFinalName = ContainerHelper::insertByUnusedName( xStylesNC, Any( xStyle ), aStyleName, ' ', bBuiltin );
                    }

                    // write style formatting properties
                    PropertySet aPropSet( xStyle );
                    getStyles().writeStyleXfToPropertySet( aPropSet, nXfId );
                }
                catch( Exception& )
                {
                    OSL_ENSURE( false, "XfCellStyle::createStyleSheet - cannot create style sheet" );
                }
            }
        }
    }
    return maFinalName;
}

// ============================================================================

namespace {

sal_Int32 lclTintToColor( sal_Int32 nColor, double fTint )
{
    if( nColor == 0x000000 )
        return 0x010101 * static_cast< sal_Int32 >( ::std::max( fTint, 0.0 ) * 255.0 );
    if( nColor == 0xFFFFFF )
        return 0x010101 * static_cast< sal_Int32 >( ::std::min( fTint + 1.0, 1.0 ) * 255.0 );

    sal_Int32 nR = (nColor >> 16) & 0xFF;
    sal_Int32 nG = (nColor >> 8) & 0xFF;
    sal_Int32 nB = nColor & 0xFF;

    double fMean = (::std::min( ::std::min( nR, nG ), nB ) + ::std::max( ::std::max( nR, nG ), nB )) / 2.0;
    double fTintTh = (fMean <= 127.5) ? ((127.5 - fMean) / (255.0 - fMean)) : (127.5 / fMean - 1.0);
    if( (fTintTh < 0.0) || ((fTintTh == 0.0) && (fTint <= 0.0)) )
    {
        double fTintMax = 255.0 / fMean - 1.0;
        double fRTh = fTintTh / fTintMax * (255.0 - nR) + nR;
        double fGTh = fTintTh / fTintMax * (255.0 - nG) + nG;
        double fBTh = fTintTh / fTintMax * (255.0 - nB) + nB;
        if( fTint <= fTintTh )
        {
            double fFactor = (fTint + 1.0) / (fTintTh + 1.0);
            nR = static_cast< sal_Int32 >( fFactor * fRTh + 0.5 );
            nG = static_cast< sal_Int32 >( fFactor * fGTh + 0.5 );
            nB = static_cast< sal_Int32 >( fFactor * fBTh + 0.5 );
        }
        else
        {
            double fFactor = (fTint > 0.0) ? (fTint * fTintMax / fTintTh) : (fTint / fTintTh);
            nR = static_cast< sal_Int32 >( fFactor * fRTh + (1.0 - fFactor) * nR + 0.5 );
            nG = static_cast< sal_Int32 >( fFactor * fGTh + (1.0 - fFactor) * nG + 0.5 );
            nB = static_cast< sal_Int32 >( fFactor * fBTh + (1.0 - fFactor) * nB + 0.5 );
        }
    }
    else
    {
        double fTintMin = fMean / (fMean - 255.0);
        double fRTh = (1.0 - fTintTh / fTintMin) * nR;
        double fGTh = (1.0 - fTintTh / fTintMin) * nG;
        double fBTh = (1.0 - fTintTh / fTintMin) * nB;
        if( fTint <= fTintTh )
        {
            double fFactor = (fTint < 0.0) ? (fTint * -fTintMin / fTintTh) : (fTint / fTintTh);
            nR = static_cast< sal_Int32 >( fFactor * fRTh + (1.0 - fFactor) * nR + 0.5 );
            nG = static_cast< sal_Int32 >( fFactor * fGTh + (1.0 - fFactor) * nG + 0.5 );
            nB = static_cast< sal_Int32 >( fFactor * fBTh + (1.0 - fFactor) * nB + 0.5 );
        }
        else
        {
            double fFactor = (1.0 - fTint) / (1.0 - fTintTh);
            nR = static_cast< sal_Int32 >( 255.5 - fFactor * (255.0 - fRTh) );
            nG = static_cast< sal_Int32 >( 255.5 - fFactor * (255.0 - fGTh) );
            nB = static_cast< sal_Int32 >( 255.5 - fFactor * (255.0 - fBTh) );
        }
    }

    return (nR << 16) | (nG << 8) | nB;
}

} // namespace

// ----------------------------------------------------------------------------

StylesBuffer::StylesBuffer( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maPalette( rGlobalData ),
    maNumFmts( rGlobalData ),
    maDefStyleName( lclGetBuiltinStyleName( OOX_STYLE_NORMAL, OUString() ) ),
    mnDefStyleXf( -1 )
{
}

BorderRef StylesBuffer::createBorder( sal_Int32* opnBorderId )
{
    if( opnBorderId ) *opnBorderId = static_cast< sal_Int32 >( maBorders.size() );
    BorderRef xBorder( new Border( getGlobalData() ) );
    maBorders.push_back( xBorder );
    return xBorder;
}

FillRef StylesBuffer::createFill( sal_Int32* opnFillId )
{
    if( opnFillId ) *opnFillId = static_cast< sal_Int32 >( maFills.size() );
    FillRef xFill( new Fill( getGlobalData() ) );
    maFills.push_back( xFill );
    return xFill;
}

void StylesBuffer::importPaletteColor( const AttributeList& rAttribs )
{
    maPalette.appendColor( rAttribs.getHex( XML_rgb, API_RGB_TRANSPARENT ) );
}

FontRef StylesBuffer::importFont( const AttributeList& )
{
    FontRef xFont( new Font( getGlobalData() ) );
    maFonts.push_back( xFont );
    return xFont;
}

void StylesBuffer::importNumFmt( const AttributeList& rAttribs )
{
    maNumFmts.importNumFmt( rAttribs );
}

BorderRef StylesBuffer::importBorder( const AttributeList& rAttribs )
{
    BorderRef xBorder = createBorder();
    xBorder->importBorder( rAttribs );
    return xBorder;
}

FillRef StylesBuffer::importFill( const AttributeList& )
{
    return createFill();
}

XfRef StylesBuffer::importXf( sal_Int32 nContext, const AttributeList& rAttribs )
{
    XfRef xXf;
    switch( nContext )
    {
        case XLS_TOKEN( cellXfs ):
            xXf.reset( new Xf( getGlobalData() ) );
            maCellXfs.push_back( xXf );
            xXf->importCellXf( rAttribs, true );
        break;
        case XLS_TOKEN( cellStyleXfs ):
            xXf.reset( new Xf( getGlobalData() ) );
            maStyleXfs.push_back( xXf );
            xXf->importCellXf( rAttribs, false );
        break;
    }
    return xXf;
}

DxfRef StylesBuffer::importDxf( const AttributeList& rAttribs )
{
    DxfRef xDxf( new Dxf( getGlobalData() ) );
    maDxfs.push_back( xDxf );
    xDxf->importDxf( rAttribs );
    return xDxf;
}

CellStyleRef StylesBuffer::importCellStyle( const AttributeList& rAttribs )
{
    CellStyleRef xCellStyle;
    sal_Int32 nXfId = rAttribs.getInteger( XML_xfId, -1 );
    if( nXfId >= 0 )
    {
        xCellStyle.reset( new CellStyle( getGlobalData() ) );
        maCellStyles[ nXfId ] = xCellStyle;
        xCellStyle->importCellStyle( rAttribs );
        if( xCellStyle->isDefaultStyle() )
            mnDefStyleXf = nXfId;
    }
    return xCellStyle;
}

void StylesBuffer::importPalette( BiffInputStream& rStrm )
{
    maPalette.importPalette( rStrm );
}

void StylesBuffer::importFont( BiffInputStream& rStrm )
{
    /* Font with index 4 is not stored in BIFF. This means effectively, first
        font in the BIFF file has index 0, fourth font has index 3, and fifth
        font has index 5. Insert a dummy font to correctly map passed font
        identifiers. */
    if( maFonts.size() == 4 )
        maFonts.push_back( maFonts.front() );

    FontRef xFont( new Font( getGlobalData() ) );
    maFonts.push_back( xFont );
    xFont->importFont( rStrm );

    /*  #i71033# Set stream text encoding from application font, if CODEPAGE
        record is missing. Must be done now (not while finalizeImport() runs),
        to be able to read all following byte strings correctly (e.g. cell
        style names). */
    if( maFonts.size() == 1 )
        setAppFontEncoding( xFont->getFontEncoding() );
}

void StylesBuffer::importFontColor( BiffInputStream& rStrm )
{
    if( !maFonts.empty() )
        maFonts.back()->importFontColor( rStrm );
}

void StylesBuffer::importFormat( BiffInputStream& rStrm )
{
    maNumFmts.importFormat( rStrm );
}

void StylesBuffer::importXf( BiffInputStream& rStrm )
{
    XfRef xXf( new Xf( getGlobalData() ) );
    // store XF in both lists (except BIFF2 which does not support cell styles)
    maCellXfs.push_back( xXf );
    if( getBiff() != BIFF2 )
        maStyleXfs.push_back( xXf );
    xXf->importXf( rStrm );
}

void StylesBuffer::importStyle( BiffInputStream& rStrm )
{
    sal_uInt16 nStyleXf;
    rStrm >> nStyleXf;
    sal_Int32 nXfId = static_cast< sal_Int32 >( nStyleXf & BIFF_STYLE_XFMASK );
    CellStyleRef xCellStyle( new CellStyle( getGlobalData() ) );
    maCellStyles[ nXfId ] = xCellStyle;
    xCellStyle->importStyle( rStrm, getFlag( nStyleXf, BIFF_STYLE_BUILTIN ) );
    if( xCellStyle->isDefaultStyle() )
        mnDefStyleXf = nXfId;
}

void StylesBuffer::finalizeImport()
{
    // fonts first, are needed to finalize unit converter and XFs below
    maFonts.forEachMem( &Font::finalizeImport );
    // finalize unit converter after default font is known
    getUnitConverter().finalizeImport();
    // number formats
    maNumFmts.finalizeImport();
    // borders and fills
    maBorders.forEachMem( &Border::finalizeImport );
    maFills.forEachMem( &Fill::finalizeImport );

    /*  Style XFs and cell XFs. The BIFF format stores cell XFs and style XFs
        mixed in a single list. The import filter has stored the XFs in both
        lists to make the getStyleXf() function working correctly (e.g. for
        retrieving the default font, see getDefaultFont() function), except for
        BIFF2 which does not support cell styles at all. Therefore, if in BIFF
        filter mode, we do not need to finalize the cell styles list. */
    if( getFilterType() == FILTER_OOX )
        maStyleXfs.forEachMem( &Xf::finalizeImport );
    maCellXfs.forEachMem( &Xf::finalizeImport );

    /*  Create user-defined and modified builtin cell styles, passing the
        boolean parameter to createStyleSheet() skips unchanged builtin styles
        (but always create the default style). */
    for( CellStyleMap::iterator aIt = maCellStyles.begin(), aEnd = maCellStyles.end(); aIt != aEnd; ++aIt )
        aIt->second->createStyleSheet( aIt->first, !aIt->second->isDefaultStyle() );
}

sal_Int32 StylesBuffer::getColor( const OoxColor& rColor ) const
{
    sal_Int32 nColor = API_RGB_TRANSPARENT;
    switch( rColor.meType )
    {
        case OoxColor::TYPE_RGB:
            nColor = rColor.mnValue & 0xFFFFFF;
        break;
        case OoxColor::TYPE_THEME:
            nColor = getTheme().getColorByIndex( rColor.mnValue );
        break;
        case OoxColor::TYPE_PALETTE:
            nColor = maPalette.getColor( rColor.mnValue );
        break;
    }
    if( (nColor != API_RGB_TRANSPARENT) && (rColor.mfTint >= -1.0) && (rColor.mfTint != 0.0) && (rColor.mfTint <= 1.0) )
        nColor = lclTintToColor( nColor, rColor.mfTint );
    return nColor;
}

FontRef StylesBuffer::getFont( sal_Int32 nFontId ) const
{
    return maFonts.get( nFontId );
}

XfRef StylesBuffer::getCellXf( sal_Int32 nXfId ) const
{
    return maCellXfs.get( nXfId );
}

XfRef StylesBuffer::getStyleXf( sal_Int32 nXfId ) const
{
    return maStyleXfs.get( nXfId );
}

DxfRef StylesBuffer::getDxf( sal_Int32 nDxfId ) const
{
    return maDxfs.get( nDxfId );
}

FontRef StylesBuffer::getFontFromCellXf( sal_Int32 nXfId ) const
{
    FontRef xFont;
    if( const Xf* pXf = getCellXf( nXfId ).get() )
        xFont = pXf->getFont();
    return xFont;
}

FontRef StylesBuffer::getDefaultFont() const
{
    FontRef xDefFont;
    if( const Xf* pXf = getStyleXf( mnDefStyleXf ).get() )
        xDefFont = pXf->getFont();
    // no font from styles - try first loaded font (e.g. BIFF2)
    if( !xDefFont )
        xDefFont = maFonts.get( 0 );
    OSL_ENSURE( xDefFont.get(), "StylesBuffer::getDefaultFont - no default font found" );
    return xDefFont;
}

const OoxFontData& StylesBuffer::getDefaultFontData() const
{
    FontRef xDefFont = getStyles().getDefaultFont();
    return xDefFont.get() ? xDefFont->getFontData() : getTheme().getDefaultFontData();
}

const OUString& StylesBuffer::createStyleSheet( sal_Int32 nXfId ) const
{
    if( CellStyle* pCellStyle = maCellStyles.get( nXfId ).get() )
        return pCellStyle->createStyleSheet( nXfId );
    // on error: fallback to default style
    return maDefStyleName;
}

void StylesBuffer::writeFontToPropertySet( PropertySet& rPropSet, sal_Int32 nFontId ) const
{
    if( Font* pFont = maFonts.get( nFontId ).get() )
        pFont->writeToPropertySet( rPropSet, FONT_PROPTYPE_CELL );
}

void StylesBuffer::writeNumFmtToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nNumFmtId ) const
{
    maNumFmts.writeToPropertySet( rPropSet, nNumFmtId );
}

void StylesBuffer::writeBorderToPropertySet( PropertySet& rPropSet, sal_Int32 nBorderId ) const
{
    if( Border* pBorder = maBorders.get( nBorderId ).get() )
        pBorder->writeToPropertySet( rPropSet );
}

void StylesBuffer::writeFillToPropertySet( PropertySet& rPropSet, sal_Int32 nFillId ) const
{
    if( Fill* pFill = maFills.get( nFillId ).get() )
        pFill->writeToPropertySet( rPropSet );
}

void StylesBuffer::writeCellXfToPropertySet( PropertySet& rPropSet, sal_Int32 nXfId ) const
{
    if( Xf* pXf = maCellXfs.get( nXfId ).get() )
        pXf->writeToPropertySet( rPropSet );
}

void StylesBuffer::writeStyleXfToPropertySet( PropertySet& rPropSet, sal_Int32 nXfId ) const
{
    if( Xf* pXf = maStyleXfs.get( nXfId ).get() )
        pXf->writeToPropertySet( rPropSet );
}

// ============================================================================

} // namespace xls
} // namespace oox

