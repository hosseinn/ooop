/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sheetviewscontext.cxx,v $
 *
 *  $Revision: 1.1.2.13 $
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

#include "oox/xls/sheetviewscontext.hxx"
#include "oox/xls/viewsettings.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/addressconverter.hxx"

using ::oox::core::AttributeList;

namespace oox {
namespace xls {

// ============================================================================

OoxSheetViewsContext::OoxSheetViewsContext( const WorksheetFragmentBase& rFragment ) :
    WorksheetContextBase( rFragment ),
    mpViewData( 0 )
{
}

bool OoxSheetViewsContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sheetViews ):   return (nElement == XLS_TOKEN( sheetView ));
        case XLS_TOKEN( sheetView ):    return (nElement == XLS_TOKEN( pane )) ||
                                               (nElement == XLS_TOKEN( selection ));
    }
    return false;
}

void OoxSheetViewsContext::onStartElement( const AttributeList& rAttrib )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sheetView ):    importSheetView( rAttrib ); break;
        case XLS_TOKEN( pane ):         importPane( rAttrib );      break;
        case XLS_TOKEN( selection ):    importSelection( rAttrib ); break;
    }
}

void OoxSheetViewsContext::importSheetView( const AttributeList& rAttribs )
{
    mpViewData = &createSheetViewData();
    mpViewData->maGridColor.meType = OoxColor::TYPE_PALETTE;
    mpViewData->maGridColor.mnValue = rAttribs.getInteger( XML_colorId, OOX_COLOR_WINDOWTEXT );
    mpViewData->maFirstPos = getAddressConverter().createValidCellAddress( rAttribs.getString( XML_topLeftCell ), getSheetIndex(), false );
    mpViewData->mnWorkbookViewId = rAttribs.getToken( XML_workbookViewId, 0 );
    mpViewData->mnViewType = rAttribs.getToken( XML_view, XML_normal );
    mpViewData->mnNormalZoom = rAttribs.getInteger( XML_zoomScaleNormal, 0 );
    mpViewData->mnPageZoom = rAttribs.getInteger( XML_zoomScaleSheetLayoutView, 0 );
    mpViewData->mnCurrentZoom = rAttribs.getInteger( XML_zoomScale, 100 );
    mpViewData->mbSelected = rAttribs.getBool( XML_tabSelected, false );
    mpViewData->mbRightToLeft = rAttribs.getBool( XML_rightToLeft, false );
    mpViewData->mbDefGridColor = rAttribs.getBool( XML_defaultGridColor, true );
    mpViewData->mbShowFormulas = rAttribs.getBool( XML_showFormulas, false );
    mpViewData->mbShowGrid = rAttribs.getBool( XML_showGridLines, true );
    mpViewData->mbShowHeadings = rAttribs.getBool( XML_showRowColHeaders, true );
    mpViewData->mbShowZeros = rAttribs.getBool( XML_showZeros, true );
    mpViewData->mbShowOutline = rAttribs.getBool( XML_showOutlineSymbols, true );
}

void OoxSheetViewsContext::importPane( const AttributeList& rAttribs )
{
    if( mpViewData )
    {
        mpViewData->maSecondPos = getAddressConverter().createValidCellAddress( rAttribs.getString( XML_topLeftCell ), getSheetIndex(), false );
        mpViewData->mnActivePaneId = rAttribs.getToken( XML_activePane, XML_topLeft );
        mpViewData->mnPaneState = rAttribs.getToken( XML_state, XML_split );
        mpViewData->mfSplitX = ::std::max< double >( rAttribs.getDouble( XML_xSplit, 0.0 ), 0.0 );
        mpViewData->mfSplitY = ::std::max< double >( rAttribs.getDouble( XML_ySplit, 0.0 ), 0.0 );
    }
}

void OoxSheetViewsContext::importSelection( const AttributeList& rAttribs )
{
    if( mpViewData )
    {
        // pane this selection belongs to
        sal_Int32 nPaneId = rAttribs.getToken( XML_pane, XML_topLeft );
        OoxSheetSelectionData& rSelData = mpViewData->createSelectionData( nPaneId );
        // cursor position
        rSelData.maActiveCell = getAddressConverter().createValidCellAddress( rAttribs.getString( XML_activeCell ), getSheetIndex(), false );
        rSelData.mnActiveCellId = rAttribs.getInteger( XML_activeCellId, 0 );
        // selection
        rSelData.maSelection.clear();
        getAddressConverter().convertToCellRangeList( rSelData.maSelection, rAttribs.getString( XML_sqref ), getSheetIndex(), false );
    }
}

// ============================================================================

namespace {

sal_Int32 lclGetOoxPaneId( sal_uInt8 nBiffPaneId, sal_Int32 nDefaultPaneId )
{
    static const sal_Int32 spnPaneIds[] = { XML_bottomRight, XML_topRight, XML_bottomLeft, XML_topLeft };
    return (nBiffPaneId < STATIC_TABLE_SIZE( spnPaneIds )) ? spnPaneIds[ nBiffPaneId ] : nDefaultPaneId;
}

} // namespace

// ----------------------------------------------------------------------------

BiffSheetViewsContext::BiffSheetViewsContext( const WorksheetHelper& rSheetHelper ) :
    WorksheetHelper( rSheetHelper ),
    mrViewData( createSheetViewData() )
{
}

void BiffSheetViewsContext::importRecord( BiffInputStream& rStrm )
{
    sal_uInt16 nRecId = rStrm.getRecId();
    switch( nRecId )
    {
        // records in all BIFF versions
        case BIFF_ID_PANE:          importPane( rStrm );        break;
        case BIFF_ID_SELECTION:     importSelection( rStrm );   break;

        // BIFF specific records
        default: switch( getBiff() )
        {
            case BIFF2: switch( nRecId )
            {
                case BIFF2_ID_WINDOW2:      importWindow2Biff2( rStrm );    break;
            }
            break;

            case BIFF3:
            case BIFF4: switch( nRecId )
            {
                case BIFF3_ID_WINDOW2:      importWindow2Biff3( rStrm );    break;
            }
            break;

            case BIFF5:
            case BIFF8: switch( nRecId )
            {
                case BIFF_ID_SCL:           importScl( rStrm );             break;
                case BIFF3_ID_WINDOW2:      importWindow2Biff3( rStrm );    break;
            }
            break;

            case BIFF_UNKNOWN: break;
        }
    }
}

void BiffSheetViewsContext::importPane( BiffInputStream& rStrm )
{
    sal_uInt8 nActivePaneId;
    sal_uInt16 nSplitX, nSplitY;
    BiffAddress aSecondPos;
    rStrm >> nSplitX >> nSplitY >> aSecondPos >> nActivePaneId;

    mrViewData.mfSplitX = nSplitX;
    mrViewData.mfSplitY = nSplitY;
    mrViewData.maSecondPos = getAddressConverter().createValidCellAddress( aSecondPos, getSheetIndex(), false );
    mrViewData.mnActivePaneId = lclGetOoxPaneId( nActivePaneId, XML_topLeft );
}

void BiffSheetViewsContext::importScl( BiffInputStream& rStrm )
{
    sal_uInt16 nNum, nDenom;
    rStrm >> nNum >> nDenom;
    OSL_ENSURE( nDenom > 0, "BiffSheetViewsContext::importScl - invalid denominator" );
    if( nDenom > 0 )
        mrViewData.mnCurrentZoom = getLimitedValue< sal_Int32, sal_uInt16 >( (nNum * 100) / nDenom, 10, 400 );
}

void BiffSheetViewsContext::importSelection( BiffInputStream& rStrm )
{
    // pane this selection belongs to
    sal_uInt8 nPaneId;
    rStrm >> nPaneId;
    OoxSheetSelectionData& rSelData = mrViewData.createSelectionData( lclGetOoxPaneId( nPaneId, -1 ) );
    // cursor position
    BiffAddress aActiveCell;
    sal_uInt16 nActiveCellId;
    rStrm >> aActiveCell >> nActiveCellId;
    rSelData.maActiveCell = getAddressConverter().createValidCellAddress( aActiveCell, getSheetIndex(), false );
    rSelData.mnActiveCellId = nActiveCellId;
    // selection
    rSelData.maSelection.clear();
    BiffRangeList aBiffSelection;
    aBiffSelection.read( rStrm, false );
    getAddressConverter().convertToCellRangeList( rSelData.maSelection, aBiffSelection, getSheetIndex(), false );
}

void BiffSheetViewsContext::importWindow2Biff2( BiffInputStream& rStrm )
{
    mrViewData.mbShowFormulas = rStrm.readuInt8() != 0;
    mrViewData.mbShowGrid = rStrm.readuInt8() != 0;
    mrViewData.mbShowHeadings = rStrm.readuInt8() != 0;
    mrViewData.mnPaneState = (rStrm.readuInt8() == 0) ? XML_split : XML_frozen;
    mrViewData.mbShowZeros = rStrm.readuInt8() != 0;
    BiffAddress aFirstPos;
    rStrm >> aFirstPos;
    mrViewData.maFirstPos = getAddressConverter().createValidCellAddress( aFirstPos, getSheetIndex(), false );
    mrViewData.mbDefGridColor = rStrm.readuInt8() != 0;
    mrViewData.maGridColor.importColorRgb( rStrm );
}

void BiffSheetViewsContext::importWindow2Biff3( BiffInputStream& rStrm )
{
    sal_uInt16 nFlags;
    BiffAddress aFirstPos;
    rStrm >> nFlags >> aFirstPos;

    mrViewData.maFirstPos = getAddressConverter().createValidCellAddress( aFirstPos, getSheetIndex(), false );
    mrViewData.mnViewType = getFlagValue( nFlags, BIFF_WIN2_PAGEBREAKMODE, XML_pageBreakPreview, XML_normal );
    mrViewData.mnPaneState = getFlagValue( nFlags, BIFF_WIN2_FROZEN, XML_frozen, XML_split );
    mrViewData.mbSelected = getFlag( nFlags, BIFF_WIN2_SELECTED );
    // #i59590# real life: Excel ignores mirror flag in chart sheets
    mrViewData.mbRightToLeft = (getSheetType() != SHEETTYPE_CHART) && getFlag( nFlags, BIFF_WIN2_RIGHTTOLEFT );
    mrViewData.mbDefGridColor = getFlag( nFlags, BIFF_WIN2_DEFGRIDCOLOR );
    mrViewData.mbShowFormulas = getFlag( nFlags, BIFF_WIN2_SHOWFORMULAS );
    mrViewData.mbShowGrid = getFlag( nFlags, BIFF_WIN2_SHOWGRID );
    mrViewData.mbShowHeadings = getFlag( nFlags, BIFF_WIN2_SHOWHEADINGS );
    mrViewData.mbShowZeros = getFlag( nFlags, BIFF_WIN2_SHOWZEROS );
    mrViewData.mbShowOutline = getFlag( nFlags, BIFF_WIN2_SHOWOUTLINE );

    if( getBiff() == BIFF8 )
    {
        mrViewData.maGridColor.importColorId( rStrm );
        // zoom data not included in chart sheets
        if( rStrm.getRecLeft() >= 6 )
        {
            rStrm.ignore( 2 );
            sal_Int16 nPageZoom, nNormalZoom;
            rStrm >> nPageZoom >> nNormalZoom;
            mrViewData.mnPageZoom = nPageZoom;
            mrViewData.mnNormalZoom = nNormalZoom;
        }
    }
    else
    {
        mrViewData.maGridColor.importColorRgb( rStrm );
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

