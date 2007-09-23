/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: viewsettings.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
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

#ifndef OOX_XLS_VIEWSETTINGS_HXX
#define OOX_XLS_VIEWSETTINGS_HXX

#include <com/sun/star/table/CellAddress.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include "oox/core/containerhelper.hxx"
#include "oox/xls/stylesbuffer.hxx"

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

/** Contains all view settings for the entire document. */
struct OoxWorkbookViewData
{
    sal_Int32           mnWinX;             /// X position of the workbook window (twips).
    sal_Int32           mnWinY;             /// Y position of the workbook window (twips).
    sal_Int32           mnWinWidth;         /// Width of the workbook window (twips).
    sal_Int32           mnWinHeight;        /// Height of the workbook window (twips).
    sal_Int32           mnActiveSheet;      /// Displayed (active) sheet.
    sal_Int32           mnFirstVisSheet;    /// First visible sheet in sheet tabbar.
    sal_Int32           mnTabBarWidth;      /// Width of sheet tabbar (1/1000 of window width).
    sal_Int32           mnVisibility;       /// Visibility state of workbook window.
    bool                mbShowTabBar;       /// True = show sheet tabbar.
    bool                mbShowHorScroll;    /// True = show horizontal sheet scrollbars.
    bool                mbShowVerScroll;    /// True = show vertical sheet scrollbars.
    bool                mbMinimized;        /// True = workbook window is minimized.

    explicit            OoxWorkbookViewData();
};

// ----------------------------------------------------------------------------

/** Contains all settings for a selection in a single pane of a sheet. */
struct OoxSheetSelectionData
{
    typedef ::std::vector< ::com::sun::star::table::CellRangeAddress > RangeAddressVector;

    ::com::sun::star::table::CellAddress maActiveCell;  /// Position of active cell (cursor).
    RangeAddressVector  maSelection;                    /// Selected cell ranges.
    sal_Int32           mnActiveCellId;                 /// Index of active cell in selection list.

    explicit            OoxSheetSelectionData();
};

// ----------------------------------------------------------------------------

/** Contains all view settings for a single sheet. */
struct OoxSheetViewData
{
    typedef ::oox::core::RefMap< sal_Int32, OoxSheetSelectionData > OoxSelectionDataMap;

    OoxSelectionDataMap maSelMap;                       /// Selections of all panes.
    OoxColor            maGridColor;                    /// Grid color.
    ::com::sun::star::table::CellAddress maFirstPos;    /// First visible cell.
    ::com::sun::star::table::CellAddress maSecondPos;   /// First visible cell in additional panes.
    sal_Int32           mnWorkbookViewId;               /// Index into list of workbookView elements.
    sal_Int32           mnViewType;                     /// View type (normal, page break, layout).
    sal_Int32           mnActivePaneId;                 /// Active pane (with cell cursor).
    sal_Int32           mnPaneState;                    /// Pane state (frozen, split).
    double              mfSplitX;                       /// Split X position (twips), or number of frozen columns.
    double              mfSplitY;                       /// Split Y position (twips), or number of frozen rows.
    sal_Int32           mnNormalZoom;                   /// Zoom factor for normal view.
    sal_Int32           mnPageZoom;                     /// Zoom factor for pagebreak preview.
    sal_Int32           mnCurrentZoom;                  /// Zoom factor for current view.
    bool                mbSelected;                     /// True = sheet is selected.
    bool                mbRightToLeft;                  /// True = sheet in right-to-left mode.
    bool                mbDefGridColor;                 /// True = default grid color.
    bool                mbShowFormulas;                 /// True = show formulas instead of results.
    bool                mbShowGrid;                     /// True = show cell grid.
    bool                mbShowHeadings;                 /// True = show column/row headings.
    bool                mbShowZeros;                    /// True = show zero value zells.
    bool                mbShowOutline;                  /// True = show outlines.

    explicit            OoxSheetViewData();

    /** Returns true, if page break preview is active. */
    bool                isPageBreakPreview() const;
    /** Returns the zoom in normal view (returns default, if current value is 0). */
    sal_Int32           getNormalZoom() const;
    /** Returns the zoom in page preview (returns default, if current value is 0). */
    sal_Int32           getPageZoom() const;

    /** Returns the selection data, if available, otherwise 0. */
    const OoxSheetSelectionData* getSelectionData( sal_Int32 nPaneId ) const;
    /** Returns the selection data of the active pane. */
    const OoxSheetSelectionData* getActiveSelectionData() const;
    /** Returns read/write access to the selection data of the specified pane. */
    OoxSheetSelectionData& createSelectionData( sal_Int32 nPaneId );
};

// ============================================================================

class ViewSettings : public GlobalDataHelper
{
public:
    explicit            ViewSettings( const GlobalDataHelper& rGlobalData );

    OoxWorkbookViewData& createWorkbookViewData();

    OoxSheetViewData&   createSheetViewData( sal_Int32 nSheet );

    void                finalizeImport();

private:
    typedef ::oox::core::RefVector< OoxWorkbookViewData >       WorkbookViewDataVec;
    typedef ::oox::core::RefVector< OoxSheetViewData >          SheetViewDataVec;
    typedef ::oox::core::RefMap< sal_Int32, SheetViewDataVec >  SheetViewDataMap;

    WorkbookViewDataVec maBookDatas;
    SheetViewDataMap    maSheetDatas;
    ::rtl::OUString     maLayoutProp;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

