/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: pagestyle.cxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 13:35:32 $
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

#include "oox/xls/pagestyle.hxx"
#include <algorithm>
#include <com/sun/star/awt/Size.hpp>
#include <com/sun/star/sheet/XHeaderFooterContent.hpp>
#include "oox/core/attributelist.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/contexthelper.hxx"
#include "oox/xls/unitconverter.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::sheet::XHeaderFooterContent;
using ::oox::core::AttributeList;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

namespace {

const double OOX_MARGIN_DEFAULT_LR          = 0.748;    /// Left/right default margin in inches.
const double OOX_MARGIN_DEFAULT_TB          = 0.984;    /// Top/bottom default margin in inches.
const double OOX_MARGIN_DEFAULT_HF          = 0.512;    /// Header/footer default margin in inches.

} // namespace

// ============================================================================

OoxPageStyleData::OoxPageStyleData() :
    mfLeftMargin( OOX_MARGIN_DEFAULT_LR ),
    mfRightMargin( OOX_MARGIN_DEFAULT_LR ),
    mfTopMargin( OOX_MARGIN_DEFAULT_TB ),
    mfBottomMargin( OOX_MARGIN_DEFAULT_TB ),
    mfHeaderMargin( OOX_MARGIN_DEFAULT_HF ),
    mfFooterMargin( OOX_MARGIN_DEFAULT_HF ),
    mnPaperSize( 1 ),
    mnCopies( 1 ),
    mnScale( 100 ),
    mnFirstPage( 1 ),
    mnFitToWidth( 1 ),
    mnFitToHeight( 1 ),
    mnHorPrintRes( 600 ),
    mnVerPrintRes( 600 ),
    mnOrientation( XML_default ),
    mnPageOrder( XML_downThenOver ),
    mnCellComments( XML_none ),
    mnPrintErrors( XML_displayed ),
    mbUseEvenHF( false ),
    mbUseFirstHF( false ),
    mbUsePrinterDefs( true ),
    mbUseFirstPage( false ),
    mbBlackWhite( false ),
    mbDraftQuality( false ),
    mbFitToPages( false ),
    mbHorCenter( false ),
    mbVerCenter( false ),
    mbPrintGrid( false ),
    mbPrintHeadings( false ),
    mbValidSettings( true )
{
}

// ============================================================================

PageStyle::PageStyle( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void PageStyle::importPageMargins( const AttributeList& rAttribs )
{
    maData.mfLeftMargin   = rAttribs.getDouble( XML_left,   OOX_MARGIN_DEFAULT_LR );
    maData.mfRightMargin  = rAttribs.getDouble( XML_right,  OOX_MARGIN_DEFAULT_LR );
    maData.mfTopMargin    = rAttribs.getDouble( XML_top,    OOX_MARGIN_DEFAULT_TB );
    maData.mfBottomMargin = rAttribs.getDouble( XML_bottom, OOX_MARGIN_DEFAULT_TB );
    maData.mfHeaderMargin = rAttribs.getDouble( XML_header, OOX_MARGIN_DEFAULT_HF );
    maData.mfFooterMargin = rAttribs.getDouble( XML_footer, OOX_MARGIN_DEFAULT_HF );
}

void PageStyle::importPageSetup( const AttributeList& rAttribs )
{
    maData.mnPaperSize      = rAttribs.getInteger( XML_paperSize, 1 );
    maData.mnCopies         = rAttribs.getInteger( XML_copies, 1 );
    maData.mnScale          = rAttribs.getInteger( XML_scale, 100 );
    maData.mnFirstPage      = rAttribs.getInteger( XML_firstPageNumber, 1 );
    maData.mnFitToWidth     = rAttribs.getInteger( XML_fitToWidth, 1 );
    maData.mnFitToHeight    = rAttribs.getInteger( XML_fitToHeight, 1 );
    maData.mnHorPrintRes    = rAttribs.getInteger( XML_horizontalDpi, 600 );
    maData.mnVerPrintRes    = rAttribs.getInteger( XML_verticalDpi, 600 );
    maData.mnOrientation    = rAttribs.getToken( XML_orientation, XML_default );
    maData.mnPageOrder      = rAttribs.getToken( XML_pageOrder, XML_downThenOver );
    maData.mnCellComments   = rAttribs.getToken( XML_cellComments, XML_none );
    maData.mnPrintErrors    = rAttribs.getToken( XML_errors, XML_displayed );
    maData.mbUsePrinterDefs = rAttribs.getBool( XML_useFirstPageNumber, true );
    maData.mbUseFirstPage   = rAttribs.getBool( XML_useFirstPageNumber, false );
    maData.mbBlackWhite     = rAttribs.getBool( XML_blackAndWhite, false );
    maData.mbDraftQuality   = rAttribs.getBool( XML_draft, false );
    maData.mbValidSettings  = true;
}

void PageStyle::importPrintOptions( const AttributeList& rAttribs )
{
    maData.mbHorCenter     = rAttribs.getBool( XML_horizontalCentered, false );
    maData.mbVerCenter     = rAttribs.getBool( XML_verticalCentered, false );
    maData.mbPrintGrid     = rAttribs.getBool( XML_gridLines, false );
    maData.mbPrintHeadings = rAttribs.getBool( XML_headings, false );
}

void PageStyle::importHeaderFooter( const AttributeList& rAttribs )
{
    maData.mbUseEvenHF = rAttribs.getBool( XML_differentOddEven, false );
    maData.mbUseFirstHF = rAttribs.getBool( XML_differentFirst, false );
}

void PageStyle::importHeaderFooterCharacters( const OUString& rChars, sal_Int32 nElement )
{
    switch( nElement )
    {
        case XLS_TOKEN( oddHeader ):    maData.maOddHeader += rChars;       break;
        case XLS_TOKEN( oddFooter ):    maData.maOddFooter += rChars;       break;
        case XLS_TOKEN( evenHeader ):   maData.maEvenHeader += rChars;      break;
        case XLS_TOKEN( evenFooter ):   maData.maEvenFooter += rChars;      break;
        case XLS_TOKEN( firstHeader ):  maData.maFirstHeader += rChars;     break;
        case XLS_TOKEN( firstFooter ):  maData.maFirstFooter += rChars;     break;
    }
}

void PageStyle::importLeftMargin( BiffInputStream& rStrm )
{
    rStrm >> maData.mfLeftMargin;
}

void PageStyle::importRightMargin( BiffInputStream& rStrm )
{
    rStrm >> maData.mfRightMargin;
}

void PageStyle::importTopMargin( BiffInputStream& rStrm )
{
    rStrm >> maData.mfTopMargin;
}

void PageStyle::importBottomMargin( BiffInputStream& rStrm )
{
    rStrm >> maData.mfBottomMargin;
}

void PageStyle::importSetup( BiffInputStream& rStrm )
{
    sal_uInt16 nPaperSize, nScale, nFirstPage, nFitToWidth, nFitToHeight, nFlags;
    rStrm >> nPaperSize >> nScale >> nFirstPage >> nFitToWidth >> nFitToHeight >> nFlags;

    maData.mnPaperSize      = nPaperSize;   // equal in BIFF and OOX
    maData.mnScale          = nScale;
    maData.mnFirstPage      = nFirstPage;
    maData.mnFitToWidth     = nFitToWidth;
    maData.mnFitToHeight    = nFitToHeight;
    maData.mnOrientation    = getFlagValue( nFlags, BIFF_SETUP_PORTRAIT, XML_portrait, XML_landscape );
    maData.mnPageOrder      = getFlagValue( nFlags, BIFF_SETUP_INROWS, XML_overThenDown, XML_downThenOver );
    maData.mbUsePrinterDefs = false;
    maData.mbUseFirstPage   = true;
    maData.mbBlackWhite     = getFlag( nFlags, BIFF_SETUP_BLACKWHITE );
    maData.mbValidSettings  = !getFlag( nFlags, BIFF_SETUP_INVALIDSETTINGS );

    if( getBiff() >= BIFF5 )
    {
        sal_uInt16 nHorPrintRes, nVerPrintRes, nCopies;
        rStrm >> nHorPrintRes >> nVerPrintRes >> maData.mfHeaderMargin >> maData.mfFooterMargin >> nCopies;

        maData.mnOrientation  = getFlagValue( nFlags, BIFF_SETUP_USEDEFORIENT, XML_default, maData.mnOrientation );
        maData.mnHorPrintRes  = nHorPrintRes;
        maData.mnVerPrintRes  = nVerPrintRes;
        maData.mnCellComments = getFlagValue( nFlags, BIFF_SETUP_PRINTNOTES, getFlagValue( nFlags, BIFF_SETUP_NOTES_END, XML_atEnd, XML_asDisplayed ), XML_none );
        maData.mbUseFirstPage = getFlag( nFlags, BIFF_SETUP_USEFIRSTPAGE );
        maData.mbDraftQuality = getFlag( nFlags, BIFF_SETUP_DRAFTQUALITY );
    }
}

void PageStyle::importHorCenter( BiffInputStream& rStrm )
{
    maData.mbHorCenter = rStrm.readuInt16() != 0;
}

void PageStyle::importVerCenter( BiffInputStream& rStrm )
{
    maData.mbVerCenter = rStrm.readuInt16() != 0;
}

void PageStyle::importPrintHeaders( BiffInputStream& rStrm )
{
    maData.mbPrintHeadings = rStrm.readuInt16() != 0;
}

void PageStyle::importPrintGridLines( BiffInputStream& rStrm )
{
    maData.mbPrintGrid = rStrm.readuInt16() != 0;
}

void PageStyle::importHeader( BiffInputStream& rStrm )
{
    if( rStrm.getRecLeft() > 0 )
        maData.maOddHeader = (getBiff() == BIFF8) ? rStrm.readUniString() : rStrm.readByteString( false, getTextEncoding() );
    else
        maData.maOddHeader = OUString();
}

void PageStyle::importFooter( BiffInputStream& rStrm )
{
    if( rStrm.getRecLeft() > 0 )
        maData.maOddFooter = (getBiff() == BIFF8) ? rStrm.readUniString() : rStrm.readByteString( false, getTextEncoding() );
    else
        maData.maOddFooter = OUString();
}

void PageStyle::setFitToPagesMode( bool bFitToPages )
{
    maData.mbFitToPages = bFitToPages;
}

void PageStyle::writeToPropertySet( PropertySet& rPropSet, WorksheetType eSheetType )
{
    getPageStylePropertyHelper().writePageStyleProperties( rPropSet, maData, eSheetType );
}

// ============================================================================

namespace {

/** Property names for page style settings. */
const sal_Char* const sppcPageNames[] =
{
    "PageScale",
    "IsLandscape",
    "FirstPageNumber",
    "PrintDownFirst",
    "PrintAnnotations",
    "CenterHorizontally",
    "CenterVertically",
    "PrintGrid",
    "PrintHeaders",
    "LeftMargin",
    "RightMargin",
    "TopMargin",
    "BottomMargin",
    "HeaderIsOn",
    "HeaderIsShared",
    "HeaderIsDynamicHeight",
    "HeaderHeight",
    "HeaderBodyDistance",
    "FooterIsOn",
    "FooterIsShared",
    "FooterIsDynamicHeight",
    "FooterHeight",
    "FooterBodyDistance",
    0
};

/** Paper size in 1/100 millimeters. */
struct ApiPaperSize
{
    sal_Int32           mnWidth;
    sal_Int32           mnHeight;
};

#define IN2MM100( v )    static_cast< sal_Int32 >( (v) * 2540.0 + 0.5 )
#define MM2MM100( v )    static_cast< sal_Int32 >( (v) * 100.0 + 0.5 )

static const ApiPaperSize spPaperSizeTable[] =
{
    { 0, 0 },                                                //  0 - (undefined)
    { IN2MM100( 8.5 ),       IN2MM100( 11 )      },          //  1 - Letter paper
    { IN2MM100( 8.5 ),       IN2MM100( 11 )      },          //  2 - Letter small paper
    { IN2MM100( 11 ),        IN2MM100( 17 )      },          //  3 - Tabloid paper
    { IN2MM100( 17 ),        IN2MM100( 11 )      },          //  4 - Ledger paper
    { IN2MM100( 8.5 ),       IN2MM100( 14 )      },          //  5 - Legal paper
    { IN2MM100( 5.5 ),       IN2MM100( 8.5 )     },          //  6 - Statement paper
    { IN2MM100( 7.25 ),      IN2MM100( 10.5 )    },          //  7 - Executive paper
    { MM2MM100( 297 ),       MM2MM100( 420 )     },          //  8 - A3 paper
    { MM2MM100( 210 ),       MM2MM100( 297 )     },          //  9 - A4 paper
    { MM2MM100( 210 ),       MM2MM100( 297 )     },          // 10 - A4 small paper
    { MM2MM100( 148 ),       MM2MM100( 210 )     },          // 11 - A5 paper
    { MM2MM100( 250 ),       MM2MM100( 353 )     },          // 12 - B4 paper
    { MM2MM100( 176 ),       MM2MM100( 250 )     },          // 13 - B5 paper
    { IN2MM100( 8.5 ),       IN2MM100( 13 )      },          // 14 - Folio paper
    { MM2MM100( 215 ),       MM2MM100( 275 )     },          // 15 - Quarto paper
    { IN2MM100( 10 ),        IN2MM100( 14 )      },          // 16 - Standard paper
    { IN2MM100( 11 ),        IN2MM100( 17 )      },          // 17 - Standard paper
    { IN2MM100( 8.5 ),       IN2MM100( 11 )      },          // 18 - Note paper
    { IN2MM100( 3.875 ),     IN2MM100( 8.875 )   },          // 19 - #9 envelope
    { IN2MM100( 4.125 ),     IN2MM100( 9.5 )     },          // 20 - #10 envelope
    { IN2MM100( 4.5 ),       IN2MM100( 10.375 )  },          // 21 - #11 envelope
    { IN2MM100( 4.75 ),      IN2MM100( 11 )      },          // 22 - #12 envelope
    { IN2MM100( 5 ),         IN2MM100( 11.5 )    },          // 23 - #14 envelope
    { IN2MM100( 17 ),        IN2MM100( 22 )      },          // 24 - C paper
    { IN2MM100( 22 ),        IN2MM100( 34 )      },          // 25 - D paper
    { IN2MM100( 34 ),        IN2MM100( 44 )      },          // 26 - E paper
    { MM2MM100( 110 ),       MM2MM100( 220 )     },          // 27 - DL envelope
    { MM2MM100( 162 ),       MM2MM100( 229 )     },          // 28 - C5 envelope
    { MM2MM100( 324 ),       MM2MM100( 458 )     },          // 29 - C3 envelope
    { MM2MM100( 229 ),       MM2MM100( 324 )     },          // 30 - C4 envelope
    { MM2MM100( 114 ),       MM2MM100( 162 )     },          // 31 - C6 envelope
    { MM2MM100( 114 ),       MM2MM100( 229 )     },          // 32 - C65 envelope
    { MM2MM100( 250 ),       MM2MM100( 353 )     },          // 33 - B4 envelope
    { MM2MM100( 176 ),       MM2MM100( 250 )     },          // 34 - B5 envelope
    { MM2MM100( 176 ),       MM2MM100( 125 )     },          // 35 - B6 envelope
    { MM2MM100( 110 ),       MM2MM100( 230 )     },          // 36 - Italy envelope
    { IN2MM100( 3.875 ),     IN2MM100( 7.5 )     },          // 37 - Monarch envelope
    { IN2MM100( 3.625 ),     IN2MM100( 6.5 )     },          // 38 - 6 3/4 envelope
    { IN2MM100( 14.875 ),    IN2MM100( 11 )      },          // 39 - US standard fanfold
    { IN2MM100( 8.5 ),       IN2MM100( 12 )      },          // 40 - German standard fanfold
    { IN2MM100( 8.5 ),       IN2MM100( 13 )      },          // 41 - German legal fanfold
    { MM2MM100( 250 ),       MM2MM100( 353 )     },          // 42 - ISO B4
    { MM2MM100( 200 ),       MM2MM100( 148 )     },          // 43 - Japanese double postcard
    { IN2MM100( 9 ),         IN2MM100( 11 )      },          // 44 - Standard paper
    { IN2MM100( 10 ),        IN2MM100( 11 )      },          // 45 - Standard paper
    { IN2MM100( 15 ),        IN2MM100( 11 )      },          // 46 - Standard paper
    { MM2MM100( 220 ),       MM2MM100( 220 )     },          // 47 - Invite envelope
    { 0, 0 },                                                // 48 - (undefined)
    { 0, 0 },                                                // 49 - (undefined)
    { IN2MM100( 9.275 ),     IN2MM100( 12 )      },          // 50 - Letter extra paper
    { IN2MM100( 9.275 ),     IN2MM100( 15 )      },          // 51 - Legal extra paper
    { IN2MM100( 11.69 ),     IN2MM100( 18 )      },          // 52 - Tabloid extra paper
    { MM2MM100( 236 ),       MM2MM100( 322 )     },          // 53 - A4 extra paper
    { IN2MM100( 8.275 ),     IN2MM100( 11 )      },          // 54 - Letter transverse paper
    { MM2MM100( 210 ),       MM2MM100( 297 )     },          // 55 - A4 transverse paper
    { IN2MM100( 9.275 ),     IN2MM100( 12 )      },          // 56 - Letter extra transverse paper
    { MM2MM100( 227 ),       MM2MM100( 356 )     },          // 57 - SuperA/SuperA/A4 paper
    { MM2MM100( 305 ),       MM2MM100( 487 )     },          // 58 - SuperB/SuperB/A3 paper
    { IN2MM100( 8.5 ),       IN2MM100( 12.69 )   },          // 59 - Letter plus paper
    { MM2MM100( 210 ),       MM2MM100( 330 )     },          // 60 - A4 plus paper
    { MM2MM100( 148 ),       MM2MM100( 210 )     },          // 61 - A5 transverse paper
    { MM2MM100( 182 ),       MM2MM100( 257 )     },          // 62 - JIS B5 transverse paper
    { MM2MM100( 322 ),       MM2MM100( 445 )     },          // 63 - A3 extra paper
    { MM2MM100( 174 ),       MM2MM100( 235 )     },          // 64 - A5 extra paper
    { MM2MM100( 201 ),       MM2MM100( 276 )     },          // 65 - ISO B5 extra paper
    { MM2MM100( 420 ),       MM2MM100( 594 )     },          // 66 - A2 paper
    { MM2MM100( 297 ),       MM2MM100( 420 )     },          // 67 - A3 transverse paper
    { MM2MM100( 322 ),       MM2MM100( 445 )     }           // 68 - A3 extra transverse paper
};

} // namespace

// ----------------------------------------------------------------------------

PageStylePropertyHelper::HFHelperData::HFHelperData( const OUString& rLeftProp, const OUString& rRightProp ) :
    maLeftProp( rLeftProp ),
    maRightProp( rRightProp ),
    mnHeight( 0 ),
    mnBodyDist( 0 ),
    mbHasContent( false ),
    mbShareOddEven( false ),
    mbDynamicHeight( false )
{
}

// ----------------------------------------------------------------------------

PageStylePropertyHelper::PageStylePropertyHelper( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maHFParser( rGlobalData ),
    maPageProps( sppcPageNames ),
    maHeaderData( CREATE_OUSTRING( "LeftPageHeaderContent" ), CREATE_OUSTRING( "RightPageHeaderContent" ) ),
    maFooterData( CREATE_OUSTRING( "LeftPageFooterContent" ), CREATE_OUSTRING( "RightPageFooterContent" ) ),
    maSizeProp( CREATE_OUSTRING( "Size" ) ),
    maScaleToXProp( CREATE_OUSTRING( "ScaleToPagesX" ) ),
    maScaleToYProp( CREATE_OUSTRING( "ScaleToPagesY" ) )
{
}

void PageStylePropertyHelper::writePageStyleProperties(
        PropertySet& rPropSet, const OoxPageStyleData& rData, WorksheetType eSheetType )
{
    // scaling
    sal_Int16 nScale = 100;
    // scale may be 0 which indicates uninitialized
    if( rData.mbValidSettings && (rData.mnScale > 0) )
        nScale = getLimitedValue< sal_Int16, sal_Int32 >( rData.mnScale, 10, 400 );

    // paper orientation
    bool bLandscape = rData.mnOrientation == XML_landscape;
    // default orientation for current sheet type (chart sheets default to landscape)
    if( !rData.mbValidSettings || (rData.mnOrientation == XML_default) ) switch( eSheetType )
    {
        case SHEETTYPE_WORKSHEET:   bLandscape = false; break;
        case SHEETTYPE_CHART:       bLandscape = true;  break;
        case SHEETTYPE_MACRO:       bLandscape = false; break;
    }

    // paper size
    if( rData.mbValidSettings && (0 < rData.mnPaperSize) && (rData.mnPaperSize < static_cast< sal_Int32 >( STATIC_TABLE_SIZE( spPaperSizeTable ) )) )
    {
        const ApiPaperSize& rPaperSize = spPaperSizeTable[ rData.mnPaperSize ];
        ::com::sun::star::awt::Size aSize( rPaperSize.mnWidth, rPaperSize.mnHeight );
        if( bLandscape )
            ::std::swap( aSize.Width, aSize.Height );
        rPropSet.setProperty( maSizeProp, aSize );
    }

    // fit to number of pages
    if( rData.mbFitToPages )
    {
        rPropSet.setProperty( maScaleToXProp, getLimitedValue< sal_Int16, sal_Int32 >( rData.mnFitToWidth, 0, 1000 ) );
        rPropSet.setProperty( maScaleToYProp, getLimitedValue< sal_Int16, sal_Int32 >( rData.mnFitToHeight, 0, 1000 ) );
    }

    // header/footer
    convertHeaderFooterData( rPropSet, maHeaderData, rData.maOddHeader, rData.maEvenHeader, rData.mbUseEvenHF, rData.mfTopMargin,    rData.mfHeaderMargin );
    convertHeaderFooterData( rPropSet, maFooterData, rData.maOddFooter, rData.maEvenFooter, rData.mbUseEvenHF, rData.mfBottomMargin, rData.mfFooterMargin );

    // write all properties to property set
    const UnitConverter& rUnitConv = getUnitConverter();
    maPageProps
        << nScale
        << bLandscape
        << getLimitedValue< sal_Int16, sal_Int32 >( rData.mbUseFirstPage ? rData.mnFirstPage : 0, 0, 9999 )
        << (rData.mnPageOrder == XML_downThenOver)
        << (rData.mnCellComments == XML_asDisplayed)
        << rData.mbHorCenter
        << rData.mbVerCenter
        << rData.mbPrintGrid
        << rData.mbPrintHeadings
        << rUnitConv.calcMm100FromInches( rData.mfLeftMargin )
        << rUnitConv.calcMm100FromInches( rData.mfRightMargin )
        // #i23296# In Calc, "TopMargin" property is distance to top of header if enabled
        << rUnitConv.calcMm100FromInches( maHeaderData.mbHasContent ? rData.mfHeaderMargin : rData.mfTopMargin )
        // #i23296# In Calc, "BottomMargin" property is distance to bottom of footer if enabled
        << rUnitConv.calcMm100FromInches( maFooterData.mbHasContent ? rData.mfFooterMargin : rData.mfBottomMargin )
        << maHeaderData.mbHasContent
        << maHeaderData.mbShareOddEven
        << maHeaderData.mbDynamicHeight
        << maHeaderData.mnHeight
        << maHeaderData.mnBodyDist
        << maFooterData.mbHasContent
        << maFooterData.mbShareOddEven
        << maFooterData.mbDynamicHeight
        << maFooterData.mnHeight
        << maFooterData.mnBodyDist
        >> rPropSet;
}

void PageStylePropertyHelper::convertHeaderFooterData(
        PropertySet& rPropSet, HFHelperData& rHFData,
        const OUString rOddContent, const OUString rEvenContent, bool bUseEvenContent,
        double fPageMargin, double fContentMargin )
{
    bool bHasOddContent  = rOddContent.getLength() > 0;
    bool bHasEvenContent = bUseEvenContent && (rEvenContent.getLength() > 0);

    sal_Int32 nOddHeight  = bHasOddContent  ? writeHeaderFooter( rPropSet, rHFData.maRightProp, rOddContent  ) : 0;
    sal_Int32 nEvenHeight = bHasEvenContent ? writeHeaderFooter( rPropSet, rHFData.maLeftProp,  rEvenContent ) : 0;

    rHFData.mnHeight = 750;
    rHFData.mnBodyDist = 250;
    rHFData.mbHasContent = bHasOddContent || bHasEvenContent;
    rHFData.mbShareOddEven = !bUseEvenContent;
    rHFData.mbDynamicHeight = true;

    if( rHFData.mbHasContent )
    {
        // use maximum height of odd/even header/footer
        rHFData.mnHeight = ::std::max( nOddHeight, nEvenHeight );
        /*  Calc contains distance between bottom of header and top of page
            body in "HeaderBodyDistance" property, and distance between bottom
            of page body and top of footer in "FooterBodyDistance" property */
        rHFData.mnBodyDist = getUnitConverter().calcMm100FromInches( fPageMargin - fContentMargin ) - rHFData.mnHeight;
        /*  #i23296# Distance less than 0 means, header or footer overlays page
            body. As this is not possible in Calc, set fixed header or footer
            height (crop header/footer) to get correct top position of page body. */
        rHFData.mbDynamicHeight = rHFData.mnBodyDist >= 0;
        /*  "HeaderHeight" property is in fact distance from top of header to
            top of page body (including "HeaderBodyDistance").
            "FooterHeight" property is in fact distance from bottom of page
            body to bottom of footer (including "FooterBodyDistance"). */
        rHFData.mnHeight += rHFData.mnBodyDist;
        // negative body distance not allowed
        rHFData.mnBodyDist = ::std::max< sal_Int32 >( rHFData.mnBodyDist, 0 );
    }
}

sal_Int32 PageStylePropertyHelper::writeHeaderFooter(
        PropertySet& rPropSet, const OUString& rPropName, const OUString& rContent )
{
    OSL_ENSURE( rContent.getLength() > 0, "PageStylePropertyHelper::writeHeaderFooter - empty h/f string found" );
    sal_Int32 nHeight = 0;
    if( rContent.getLength() > 0 )
    {
        Reference< XHeaderFooterContent > xHFContent;
        if( rPropSet.getProperty( xHFContent, rPropName ) && xHFContent.is() )
        {
            maHFParser.parse( xHFContent, rContent );
            rPropSet.setProperty( rPropName, xHFContent );
            nHeight = getUnitConverter().calcMm100FromPoints( maHFParser.getTotalHeight() );
        }
    }
    return nHeight;
}

// ============================================================================

} // namespace xls
} // namespace oox

