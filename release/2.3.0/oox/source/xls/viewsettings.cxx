/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: viewsettings.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 12:31:23 $
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

#include "oox/xls/viewsettings.hxx"
#include <com/sun/star/container/XNameContainer.hpp>
#include <com/sun/star/container/XIndexContainer.hpp>
#include <com/sun/star/document/XViewDataSupplier.hpp>
#include <com/sun/star/text/WritingMode2.hpp>
#include "tokens.hxx"
#include "oox/core/containerhelper.hxx"
#include "oox/core/propertysequence.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/unitconverter.hxx"
#include "oox/xls/worksheetbuffer.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::container::XNameContainer;
using ::com::sun::star::container::XIndexContainer;
using ::com::sun::star::container::XIndexAccess;
using ::com::sun::star::document::XViewDataSupplier;
using ::com::sun::star::table::CellAddress;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySequence;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

namespace {

const sal_Int32 OOX_BOOKVIEW_TABBARRATIO_DEF    = 600;      /// Default tabbar ratio.
const sal_Int32 OOX_SHEETVIEW_NORMALZOOM_DEF    = 100;      /// Default zoom for normal view.
const sal_Int32 OOX_SHEETVIEW_PAGEZOOM_DEF      = 60;       /// Default zoom for pagebreak preview.

// Attention: view settings in Calc do not use com.sun.star.view.DocumentZoomType!
const sal_Int16 API_ZOOMTYPE_PERCENT            = 0;        /// Zoom value in percent.

const sal_Int32 API_ZOOMVALUE_MIN               = 20;       /// Minimum zoom in Calc.
const sal_Int32 API_ZOOMVALUE_MAX               = 400;      /// Maximum zoom in Calc.

const sal_Int16 API_SPLITMODE_NONE              = 0;        /// No splits in window.
const sal_Int16 API_SPLITMODE_SPLIT             = 1;        /// Window is split.
const sal_Int16 API_SPLITMODE_FREEZE            = 2;        /// Window has frozen panes.

const sal_Int16 API_SPLITPANE_TOPLEFT           = 0;        /// Top-left, or top pane.
const sal_Int16 API_SPLITPANE_TOPRIGHT          = 1;        /// Top-right pane.
const sal_Int16 API_SPLITPANE_BOTTOMLEFT        = 2;        /// Bottom-left, bottom, left, or single pane.
const sal_Int16 API_SPLITPANE_BOTTOMRIGHT       = 3;        /// Bottom-right, or right pane.

} // namespace

// ============================================================================

OoxWorkbookViewData::OoxWorkbookViewData() :
    mnWinX( 0 ),
    mnWinY( 0 ),
    mnWinWidth( 0 ),
    mnWinHeight( 0 ),
    mnActiveSheet( 0 ),
    mnFirstVisSheet( 0 ),
    mnTabBarWidth( OOX_BOOKVIEW_TABBARRATIO_DEF ),
    mnVisibility( XML_visible ),
    mbShowTabBar( true ),
    mbShowHorScroll( true ),
    mbShowVerScroll( true ),
    mbMinimized( false )
{
}

// ----------------------------------------------------------------------------

OoxSheetSelectionData::OoxSheetSelectionData() :
    mnActiveCellId( 0 )
{
}

// ----------------------------------------------------------------------------

OoxSheetViewData::OoxSheetViewData() :
    maGridColor( OoxColor::TYPE_PALETTE, OOX_COLOR_WINDOWTEXT ),
    mnWorkbookViewId( 0 ),
    mnViewType( XML_normal ),
    mnActivePaneId( XML_topLeft ),
    mnPaneState( XML_split ),
    mfSplitX( 0.0 ),
    mfSplitY( 0.0 ),
    mnNormalZoom( 0 ),
    mnPageZoom( 0 ),
    mnCurrentZoom( 0 ), // default to mnNormalZoom or mnPageZoom
    mbSelected( false ),
    mbRightToLeft( false ),
    mbDefGridColor( true ),
    mbShowFormulas( false ),
    mbShowGrid( true ),
    mbShowHeadings( true ),
    mbShowZeros( true ),
    mbShowOutline( true )
{
}

bool OoxSheetViewData::isPageBreakPreview() const
{
    return mnViewType == XML_pageBreakPreview;
}

sal_Int32 OoxSheetViewData::getNormalZoom() const
{
    const sal_Int32& rnZoom = isPageBreakPreview() ? mnNormalZoom : mnCurrentZoom;
    sal_Int32 nZoom = (rnZoom > 0) ? rnZoom : OOX_SHEETVIEW_NORMALZOOM_DEF;
    return getLimitedValue< sal_Int32 >( nZoom, API_ZOOMVALUE_MIN, API_ZOOMVALUE_MAX );
}

sal_Int32 OoxSheetViewData::getPageZoom() const
{
    const sal_Int32& rnZoom = isPageBreakPreview() ? mnCurrentZoom : mnPageZoom;
    sal_Int32 nZoom = (rnZoom > 0) ? rnZoom : OOX_SHEETVIEW_PAGEZOOM_DEF;
    return getLimitedValue< sal_Int32 >( nZoom, API_ZOOMVALUE_MIN, API_ZOOMVALUE_MAX );
}

const OoxSheetSelectionData* OoxSheetViewData::getSelectionData( sal_Int32 nPaneId ) const
{
    return maSelMap.get( nPaneId ).get();
}

const OoxSheetSelectionData* OoxSheetViewData::getActiveSelectionData() const
{
    return getSelectionData( mnActivePaneId );
}

OoxSheetSelectionData& OoxSheetViewData::createSelectionData( sal_Int32 nPaneId )
{
    OoxSelectionDataMap::data_type& rxSelData = maSelMap[ nPaneId ];
    if( !rxSelData )
        rxSelData.reset( new OoxSheetSelectionData );
    return *rxSelData;
}

// ============================================================================

namespace {

/** Property names for document view settings. */
const sal_Char* const sppcDocNames[] =
{
    "Tables",
    "ActiveTable",
    "HasHorizontalScrollBar",
    "HasVerticalScrollBar",
    "HasSheetTabs",
    "RelativeHorizontalTabbarWidth",
    // Excel sheet properties that are document-global in Calc
    "GridColor",
    "ZoomType",
    "ZoomValue",
    "PageViewZoomValue",
    "ShowPageBreakPreview",
    "ShowFormulas",
    "ShowGrid",
    "HasColumnRowHeaders",
    "ShowZeroValues",
    "IsOutlineSymbolsSet",
    0
};

/** Property names for sheet view settings. */
const sal_Char* const sppcSheetNames[] =
{
    "TableSelected",
    "CursorPositionX",
    "CursorPositionY",
    "HorizontalSplitMode",
    "VerticalSplitMode",
    "HorizontalSplitPositionTwips",
    "VerticalSplitPositionTwips",
    "ActiveSplitRange",
    "PositionLeft",
    "PositionTop",
    "PositionRight",
    "PositionBottom",
    // Excel sheet properties that are document-global in Calc
    "GridColor",
    "ZoomType",
    "ZoomValue",
    "PageViewZoomValue",
    "ShowPageBreakPreview",
    "ShowFormulas",
    "ShowGrid",
    "HasColumnRowHeaders",
    "ShowZeroValues",
    "IsOutlineSymbolsSet",
    0
};

} // namespace

// ----------------------------------------------------------------------------

ViewSettings::ViewSettings( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maLayoutProp( CREATE_OUSTRING( "TableLayout" ) )
{
}

OoxWorkbookViewData& ViewSettings::createWorkbookViewData()
{
    maBookDatas.push_back( WorkbookViewDataVec::value_type( new OoxWorkbookViewData ) );
    return *maBookDatas.back();
}

OoxSheetViewData& ViewSettings::createSheetViewData( sal_Int32 nSheet )
{
    SheetViewDataMap::data_type& rxData = maSheetDatas[ nSheet ];
    if( !rxData )
        rxData.reset( new SheetViewDataVec );
    rxData->push_back( SheetViewDataVec::value_type( new OoxSheetViewData ) );
    return *rxData->back();
}

void ViewSettings::finalizeImport()
{
    const WorksheetBuffer& rWorksheets = getWorksheets();
    const StylesBuffer& rStyles = getStyles();
    sal_Int32 nSheetCount = rWorksheets.getInternalSheetCount();

    // force creation of workbook view data to get the Excel defaults
    const OoxWorkbookViewData& rBookData = maBookDatas.empty() ? createWorkbookViewData() : *maBookDatas.front();
    sal_Int32 nActiveSheet = getLimitedValue< sal_Int32, sal_Int32 >( rBookData.mnActiveSheet, 0, nSheetCount - 1 );

    // view settings for all sheets
    Reference< XNameContainer > xSheetsNC = ContainerHelper::createNameContainer();
    if( !xSheetsNC.is() )
        return;

    PropertySequence aSheetProps( sppcSheetNames );
    for( sal_Int32 nSheet = 0; nSheet < nSheetCount; ++nSheet )
    {
        // force creation of sheet view data to get the Excel defaults
        SheetViewDataVec* pSheetDataVec = maSheetDatas.get( nSheet ).get();
        const OoxSheetViewData& rSheetData = (!pSheetDataVec || pSheetDataVec->empty()) ?
            createSheetViewData( nSheet ) : *pSheetDataVec->front();

        // mirrored sheet (this is not a view setting in Calc)
        if( rSheetData.mbRightToLeft )
        {
            PropertySet aPropSet( getSheet( nSheet ) );
            aPropSet.setProperty( maLayoutProp, ::com::sun::star::text::WritingMode2::RL_TB );
        }

        // current cursor position (selection not supported via API)
        const OoxSheetSelectionData* pSelData = rSheetData.getActiveSelectionData();
        CellAddress aCursor = pSelData ? pSelData->maActiveCell : rSheetData.maFirstPos;

        // freeze/split position
        sal_Int16 nHSplitMode = API_SPLITMODE_NONE;
        sal_Int16 nVSplitMode = API_SPLITMODE_NONE;
        sal_Int32 nHSplitPos = 0;
        sal_Int32 nVSplitPos = 0;
        if( rSheetData.mnPaneState == XML_frozen )
        {
            /*  Frozen panes: handle split position as row/column positions.
                #i35812# Excel uses number of visible rows/columns in the
                    frozen area (rows/columns scolled outside are not incuded),
                    Calc uses absolute position of first unfrozen row/column. */
            const CellAddress& rMaxApiPos = getAddressConverter().getMaxApiAddress();
            if( (rSheetData.mfSplitX >= 1.0) && (rSheetData.maFirstPos.Column + rSheetData.mfSplitX <= rMaxApiPos.Column) )
                nHSplitPos = static_cast< sal_Int32 >( rSheetData.maFirstPos.Column + rSheetData.mfSplitX );
            nHSplitMode = (nHSplitPos > 0) ? API_SPLITMODE_FREEZE : API_SPLITMODE_NONE;
            if( (rSheetData.mfSplitY >= 1.0) && (rSheetData.maFirstPos.Row + rSheetData.mfSplitY <= rMaxApiPos.Row) )
                nVSplitPos = static_cast< sal_Int32 >( rSheetData.maFirstPos.Row + rSheetData.mfSplitY );
            nVSplitMode = (nVSplitPos > 0) ? API_SPLITMODE_FREEZE : API_SPLITMODE_NONE;
        }
        else if( rSheetData.mnPaneState == XML_split )
        {
            // split window: view settings API uses twips...
            nHSplitPos = getLimitedValue< sal_Int32, double >( rSheetData.mfSplitX + 0.5, 0, SAL_MAX_INT32 );
            nHSplitMode = (nHSplitPos > 0) ? API_SPLITMODE_SPLIT : API_SPLITMODE_NONE;
            nVSplitPos = getLimitedValue< sal_Int32, double >( rSheetData.mfSplitY + 0.5, 0, SAL_MAX_INT32 );
            nVSplitMode = (nVSplitPos > 0) ? API_SPLITMODE_SPLIT : API_SPLITMODE_NONE;
        }

        // active pane
        sal_Int16 nActivePane = API_SPLITPANE_BOTTOMLEFT;
        switch( rSheetData.mnActivePaneId )
        {
            // no horizontal split -> always use left panes
            // no vertical split -> always use *bottom* panes
            case XML_topLeft:
                nActivePane = (nVSplitMode == API_SPLITMODE_NONE) ? API_SPLITPANE_BOTTOMLEFT : API_SPLITPANE_TOPLEFT;
            break;
            case XML_topRight:
                nActivePane = (nHSplitMode == API_SPLITMODE_NONE) ?
                    ((nVSplitMode == API_SPLITMODE_NONE) ? API_SPLITPANE_BOTTOMLEFT : API_SPLITPANE_TOPLEFT) :
                    ((nVSplitMode == API_SPLITMODE_NONE) ? API_SPLITPANE_BOTTOMRIGHT : API_SPLITPANE_TOPRIGHT);
            break;
            case XML_bottomLeft:
                nActivePane = API_SPLITPANE_BOTTOMLEFT;
            break;
            case XML_bottomRight:
                nActivePane = (nHSplitMode == API_SPLITMODE_NONE) ? API_SPLITPANE_BOTTOMLEFT : API_SPLITPANE_BOTTOMRIGHT;
            break;
        }

        // write the sheet view settings into the property sequence
        aSheetProps
            << ((nSheet == nActiveSheet) || rSheetData.mbSelected)
            << aCursor.Column
            << aCursor.Row
            << nHSplitMode
            << nVSplitMode
            << nHSplitPos
            << nVSplitPos
            << nActivePane
            << rSheetData.maFirstPos.Column
            << rSheetData.maFirstPos.Row
            << rSheetData.maSecondPos.Column
            << ((nVSplitPos > 0) ? rSheetData.maSecondPos.Row : rSheetData.maFirstPos.Row)
            << (rSheetData.mbDefGridColor ? API_RGB_TRANSPARENT : rStyles.getColor( rSheetData.maGridColor ))
            << API_ZOOMTYPE_PERCENT
            << static_cast< sal_Int16 >( rSheetData.getNormalZoom() )
            << static_cast< sal_Int16 >( rSheetData.getPageZoom() )
            << rSheetData.isPageBreakPreview()
            << rSheetData.mbShowFormulas
            << rSheetData.mbShowGrid
            << rSheetData.mbShowHeadings
            << rSheetData.mbShowZeros
            << rSheetData.mbShowOutline;

        // add sheet view settings to name container
        ContainerHelper::insertByName( xSheetsNC,
            rWorksheets.getFinalSheetName( nSheet ),
            Any( aSheetProps.createPropertySequence() ) );
    }

    // use data of active sheet to set sheet properties that are document-global in Calc
    const OoxSheetViewData& rActiveSheetData = *maSheetDatas.get( nActiveSheet )->front();

    PropertySequence aDocProps( sppcDocNames );
    aDocProps
        << xSheetsNC
        << rWorksheets.getFinalSheetName( nActiveSheet )
        << rBookData.mbShowHorScroll
        << rBookData.mbShowVerScroll
        << rBookData.mbShowTabBar
        << double( rBookData.mnTabBarWidth / 1000.0 )
        << (rActiveSheetData.mbDefGridColor ? API_RGB_TRANSPARENT : rStyles.getColor( rActiveSheetData.maGridColor ))
        << API_ZOOMTYPE_PERCENT
        << static_cast< sal_Int16 >( rActiveSheetData.getNormalZoom() )
        << static_cast< sal_Int16 >( rActiveSheetData.getPageZoom() )
        << rActiveSheetData.isPageBreakPreview()
        << rActiveSheetData.mbShowFormulas
        << rActiveSheetData.mbShowGrid
        << rActiveSheetData.mbShowHeadings
        << rActiveSheetData.mbShowZeros
        << rActiveSheetData.mbShowOutline;

    Reference< XIndexContainer > xContainer = ContainerHelper::createIndexContainer();
    if( xContainer.is() ) try
    {
        xContainer->insertByIndex( 0, Any( aDocProps.createPropertySequence() ) );
        Reference< XIndexAccess > xIAccess( xContainer, UNO_QUERY_THROW );
        Reference< XViewDataSupplier > xViewDataSuppl( getDocument(), UNO_QUERY_THROW );
        xViewDataSuppl->setViewData( xIAccess );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "ViewSettings::finalizeImport - cannot create document view settings" );
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

