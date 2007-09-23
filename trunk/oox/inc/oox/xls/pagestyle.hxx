/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: pagestyle.hxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 13:35:18 $
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

#ifndef OOX_XLS_PAGESTYLE_HXX
#define OOX_XLS_PAGESTYLE_HXX

#include "oox/core/propertysequence.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/headerfooterparser.hxx"
#include "oox/xls/worksheethelper.hxx"

namespace oox { namespace core {
    class AttributeList;
    class PropertySet;
} }

namespace oox {
namespace xls {

// ============================================================================

// record constants -----------------------------------------------------------

const sal_uInt16 BIFF_SETUP_INROWS          = 0x0001;
const sal_uInt16 BIFF_SETUP_PORTRAIT        = 0x0002;
const sal_uInt16 BIFF_SETUP_INVALIDSETTINGS = 0x0004;
const sal_uInt16 BIFF_SETUP_BLACKWHITE      = 0x0008;
const sal_uInt16 BIFF_SETUP_DRAFTQUALITY    = 0x0010;
const sal_uInt16 BIFF_SETUP_PRINTNOTES      = 0x0020;
const sal_uInt16 BIFF_SETUP_USEDEFORIENT    = 0x0040;
const sal_uInt16 BIFF_SETUP_USEFIRSTPAGE    = 0x0080;
const sal_uInt16 BIFF_SETUP_NOTES_END       = 0x0200;

// ============================================================================

/** Holds page style data for a single sheet. */
struct OoxPageStyleData
{
    ::rtl::OUString     maOddHeader;            /// Header string for odd pages.
    ::rtl::OUString     maOddFooter;            /// Footer string for odd pages.
    ::rtl::OUString     maEvenHeader;           /// Header string for even pages.
    ::rtl::OUString     maEvenFooter;           /// Footer string for even pages.
    ::rtl::OUString     maFirstHeader;          /// Header string for first page of the sheet.
    ::rtl::OUString     maFirstFooter;          /// Footer string for first page of the sheet.
    double              mfLeftMargin;           /// Margin between left edge of page and begin of sheet area.
    double              mfRightMargin;          /// Margin between end of sheet area and right edge of page.
    double              mfTopMargin;            /// Margin between top egde of page and begin of sheet area.
    double              mfBottomMargin;         /// Margin between end of sheet area and bottom edge of page.
    double              mfHeaderMargin;         /// Margin between top edge of page and begin of header.
    double              mfFooterMargin;         /// Margin between end of footer and bottom edge of page.
    sal_Int32           mnPaperSize;            /// Paper size (enumeration).
    sal_Int32           mnCopies;               /// Number of copies to print.
    sal_Int32           mnScale;                /// Page scale (zoom in percent).
    sal_Int32           mnFirstPage;            /// First page number.
    sal_Int32           mnFitToWidth;           /// Fit to number of pages in horizontal direction.
    sal_Int32           mnFitToHeight;          /// Fit to number of pages in vertical direction.
    sal_Int32           mnHorPrintRes;          /// Horizontal printing resolution in DPI.
    sal_Int32           mnVerPrintRes;          /// Vertical printing resolution in DPI.
    sal_Int32           mnOrientation;          /// Landscape or portrait.
    sal_Int32           mnPageOrder;            /// Page order through sheet area (to left or down).
    sal_Int32           mnCellComments;         /// Cell comments printing mode.
    sal_Int32           mnPrintErrors;          /// Cell error printing mode.
    bool                mbUseEvenHF;            /// True = use maEvenHeader/maEvenFooter.
    bool                mbUseFirstHF;           /// True = use maFirstHeader/maFirstFooter.
    bool                mbUsePrinterDefs;       /// True = use printer defaults; false = use schema defaults.
    bool                mbUseFirstPage;         /// True = start page numbering with mnFirstPage.
    bool                mbBlackWhite;           /// True = print black and white.
    bool                mbDraftQuality;         /// True = print in draft quality.
    bool                mbFitToPages;           /// True = Fit to width/height; false = scale in percent.
    bool                mbHorCenter;            /// True = horizontally centered.
    bool                mbVerCenter;            /// True = vertically centered.
    bool                mbPrintGrid;            /// True = print grid lines.
    bool                mbPrintHeadings;        /// True = print column/row headings.
    bool                mbValidSettings;        /// False = some of the values are not valid (BIFF only).

    explicit            OoxPageStyleData();
};

// ============================================================================

class PageStyle : public GlobalDataHelper
{
public:
    explicit            PageStyle( const GlobalDataHelper& rGlobalData );

    /** Imports pageMarings context. */
    void                importPageMargins( const ::oox::core::AttributeList& rAttribs );
    /** Imports pageSetup context. */
    void                importPageSetup( const ::oox::core::AttributeList& rAttribs );
    /** Imports printing options from a printOptions element. */
    void                importPrintOptions( const ::oox::core::AttributeList& rAttribs );
    /** Imports header and footer settings from a headerFooter element. */
    void                importHeaderFooter( const ::oox::core::AttributeList& rAttribs );
    /** Imports header/footer characters from a headerFooter element. */
    void                importHeaderFooterCharacters( const ::rtl::OUString& rChars, sal_Int32 nElement );

    /** Imports the LEFTMARGIN record from the passed BIFF stream. */
    void                importLeftMargin( BiffInputStream& rStrm );
    /** Imports the RIGHTMARGIN record from the passed BIFF stream. */
    void                importRightMargin( BiffInputStream& rStrm );
    /** Imports the TOPMARGIN record from the passed BIFF stream. */
    void                importTopMargin( BiffInputStream& rStrm );
    /** Imports the BOTTOMMARGIN record from the passed BIFF stream. */
    void                importBottomMargin( BiffInputStream& rStrm );
    /** Imports the SETUP record from the passed BIFF stream. */
    void                importSetup( BiffInputStream& rStrm );
    /** Imports the HCENTER record from the passed BIFF stream. */
    void                importHorCenter( BiffInputStream& rStrm );
    /** Imports the VCENTER record from the passed BIFF stream. */
    void                importVerCenter( BiffInputStream& rStrm );
    /** Imports the PRINTHEADERS record from the passed BIFF stream. */
    void                importPrintHeaders( BiffInputStream& rStrm );
    /** Imports the PRINTGRIDLINES record from the passed BIFF stream. */
    void                importPrintGridLines( BiffInputStream& rStrm );
    /** Imports the HEADER record from the passed BIFF stream. */
    void                importHeader( BiffInputStream& rStrm );
    /** Imports the FOOTER record from the passed BIFF stream. */
    void                importFooter( BiffInputStream& rStrm );

    /** Sets whether percentual scaling or fit to width/height scaling is used. */
    void                setFitToPagesMode( bool bFitToPages );

    /** Populates the passed property set with collected page style properties. */
    void                writeToPropertySet( ::oox::core::PropertySet& rPropSet, WorksheetType eSheetType );

private:
    OoxPageStyleData    maData;
};

// ============================================================================

class PageStylePropertyHelper : public GlobalDataHelper
{
public:
    explicit            PageStylePropertyHelper( const GlobalDataHelper& rGlobalData );

    /** Writes all properties to the passed property set of a page style object. */
    void                writePageStyleProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const OoxPageStyleData& rData,
                            WorksheetType eSheetType );

private:
    struct HFHelperData
    {
        ::rtl::OUString     maLeftProp;
        ::rtl::OUString     maRightProp;
        sal_Int32           mnHeight;
        sal_Int32           mnBodyDist;
        bool                mbHasContent;
        bool                mbShareOddEven;
        bool                mbDynamicHeight;

        explicit            HFHelperData( const ::rtl::OUString& rLeftProp, const ::rtl::OUString& rRightProp );
    };

private:
    void                convertHeaderFooterData(
                            ::oox::core::PropertySet& rPropSet,
                            HFHelperData& rHFData,
                            const ::rtl::OUString rOddContent,
                            const ::rtl::OUString rEvenContent,
                            bool bUseEvenContent,
                            double fPageMargin,
                            double fContentMargin );

    sal_Int32           writeHeaderFooter(
                            ::oox::core::PropertySet& rPropSet,
                            const ::rtl::OUString& rPropName,
                            const ::rtl::OUString& rContent );

private:
    HeaderFooterParser  maHFParser;
    ::oox::core::PropertySequence maPageProps;
    HFHelperData        maHeaderData;
    HFHelperData        maFooterData;
    ::rtl::OUString     maSizeProp;
    ::rtl::OUString     maScaleToXProp;
    ::rtl::OUString     maScaleToYProp;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

