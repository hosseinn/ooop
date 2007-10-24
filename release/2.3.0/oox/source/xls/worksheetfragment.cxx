/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: worksheetfragment.cxx,v $
 *
 *  $Revision: 1.1.2.59 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:49 $
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

#include "oox/xls/worksheetfragment.hxx"
#include <vector>
#include "oox/core/relations.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/autofiltercontext.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/condformatcontext.hxx"
#include "oox/xls/datavalidationscontext.hxx"
#include "oox/xls/externallinkbuffer.hxx"
#include "oox/xls/pagestyle.hxx"
#include "oox/xls/pivottablefragment.hxx"
#include "oox/xls/querytablefragment.hxx"
#include "oox/xls/sheetdatacontext.hxx"
#include "oox/xls/sheetviewscontext.hxx"
#include "oox/xls/workbookfragment.hxx"

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::oox::core::AttributeList;
using ::oox::core::Relation;
using ::oox::core::RelationPtr;

namespace oox {
namespace xls {

// ============================================================================

OoxWorksheetFragment::OoxWorksheetFragment( const GlobalDataHelper& rGlobalData,
        const OUString& rFragmentPath, const Reference< XSpreadsheet >& rxSheet, sal_Int32 nSheet ) :
    WorksheetFragmentBase( WorksheetHelper( rGlobalData, SHEETTYPE_WORKSHEET, rxSheet, nSheet ), rFragmentPath )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxWorksheetFragment::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XML_ROOT_CONTEXT:
            return  (nElement == XLS_TOKEN( worksheet ));
        case XLS_TOKEN( worksheet ):
            return  (nElement == XLS_TOKEN( cols )) ||
                    (nElement == XLS_TOKEN( sheetPr )) ||
                    (nElement == XLS_TOKEN( sheetFormatPr )) ||
                    (nElement == XLS_TOKEN( sheetData )) ||
                    (nElement == XLS_TOKEN( sheetViews )) ||
                    (nElement == XLS_TOKEN( rowBreaks )) ||
                    (nElement == XLS_TOKEN( colBreaks )) ||
                    (nElement == XLS_TOKEN( hyperlinks )) ||
                    (nElement == XLS_TOKEN( autoFilter )) ||
                    (nElement == XLS_TOKEN( dataValidations )) ||
                    (nElement == XLS_TOKEN( conditionalFormatting )) ||
                    (nElement == XLS_TOKEN( pageMargins )) ||
                    (nElement == XLS_TOKEN( pageSetup )) ||
                    (nElement == XLS_TOKEN( headerFooter )) ||
                    (nElement == XLS_TOKEN( printOptions )) ||
                    (nElement == XLS_TOKEN( mergeCells )) ||
                    (nElement == XLS_TOKEN( phoneticPr ));
        case XLS_TOKEN( sheetPr ):
            return  (nElement == XLS_TOKEN( outlinePr )) ||
                    (nElement == XLS_TOKEN( pageSetUpPr ));
        case XLS_TOKEN( rowBreaks ):
            return  (nElement == XLS_TOKEN( brk ));
        case XLS_TOKEN( headerFooter ):
            return  (nElement == XLS_TOKEN( firstHeader )) ||
                    (nElement == XLS_TOKEN( firstFooter )) ||
                    (nElement == XLS_TOKEN( oddHeader )) ||
                    (nElement == XLS_TOKEN( oddFooter )) ||
                    (nElement == XLS_TOKEN( evenHeader )) ||
                    (nElement == XLS_TOKEN( evenFooter ));
        case XLS_TOKEN( colBreaks ):
            return  (nElement == XLS_TOKEN( brk ));
        case XLS_TOKEN( hyperlinks ):
            return  (nElement == XLS_TOKEN( hyperlink ));
        case XLS_TOKEN( sheetViews ):
            return  (nElement == XLS_TOKEN( sheetView ));
        case XLS_TOKEN( cols ):
            return  (nElement == XLS_TOKEN( col ));
        case XLS_TOKEN( mergeCells ):
            return  (nElement == XLS_TOKEN( mergeCell ));
    }
    return false;
}

Reference< XFastContextHandler > OoxWorksheetFragment::onCreateContext( sal_Int32 nElement, const AttributeList& /*rAttribs*/ )
{
    switch( nElement )
    {
        case XLS_TOKEN( sheetData ):
            return new OoxSheetDataContext( *this );
        case XLS_TOKEN( sheetViews ):
            return new OoxSheetViewsContext( *this );
        case XLS_TOKEN( autoFilter ):
            return new OoxAutoFilterContext( *this );
        case XLS_TOKEN( dataValidations ):
            return new OoxDataValidationsContext( *this );
        case XLS_TOKEN( conditionalFormatting ):
            return new OoxCondFormatContext( *this );
    }
    return this;
}

void OoxWorksheetFragment::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( worksheet ):
            importWorksheet( rAttribs );
        break;
        case XLS_TOKEN( col ):
            importColumn( rAttribs );
        break;
        case XLS_TOKEN( mergeCell ):
            importMergeCell( rAttribs );
        break;
        case XLS_TOKEN( sheetPr ):
            importSheetPr( rAttribs );
        break;
        case XLS_TOKEN( outlinePr ):
            importOutlinePr( rAttribs );
        break;
        case XLS_TOKEN( pageSetUpPr ):
            importPageSetUpPr( rAttribs );
        break;
        case XLS_TOKEN( sheetFormatPr ):
            importSheetFormatPr( rAttribs );
        break;
        case XLS_TOKEN( brk ):
            importBrk( rAttribs );
        break;
        case XLS_TOKEN( pageMargins ):
            getPageStyle().importPageMargins( rAttribs );
        break;
        case XLS_TOKEN( pageSetup ):
            getPageStyle().importPageSetup( rAttribs );
        break;
        case XLS_TOKEN( printOptions ):
            getPageStyle().importPrintOptions( rAttribs );
        break;
        case XLS_TOKEN( headerFooter ):
            getPageStyle().importHeaderFooter( rAttribs );
        break;
        case XLS_TOKEN( hyperlink ):
            importHyperlink( rAttribs );
        break;
        case XLS_TOKEN( phoneticPr ):
            getPhoneticSettings().importPhoneticPr( rAttribs );
        break;
    }
}

void OoxWorksheetFragment::onEndElement( const OUString& rChars )
{
    sal_Int32 nCurrentContext = getCurrentContext();
    switch( nCurrentContext )
    {
        case XLS_TOKEN( firstHeader ):
        case XLS_TOKEN( firstFooter ):
        case XLS_TOKEN( oddHeader ):
        case XLS_TOKEN( oddFooter ):
        case XLS_TOKEN( evenHeader ):
        case XLS_TOKEN( evenFooter ):
            getPageStyle().importHeaderFooterCharacters( rChars, nCurrentContext );
        break;
        case XLS_TOKEN( worksheet ):
            finalizeWorksheetImport();
        break;
    }
}

// private --------------------------------------------------------------------

void OoxWorksheetFragment::importWorksheet( const AttributeList& /*rAttribs*/ )
{
    static const OUString sQTableRelTypeName = CREATE_RELATIONS_TYPE( "queryTable" );
    static const OUString sPivotTypeName     = CREATE_RELATIONS_TYPE( "pivotTable" );

    ::std::vector<RelationPtr>::const_iterator itr = getRelations()->begin(), itrEnd = getRelations()->end();
    for ( ; itr != itrEnd; ++itr )
    {
        const OUString& aType = itr->get()->msType;
        OUString aFragPath = resolveRelativePath( itr->get()->msTarget );

        Reference< XFastDocumentHandler > xHandler;
        if( aType.equals( sQTableRelTypeName )  )
            xHandler.set( new OoxQueryTableFragment( getGlobalData(), aFragPath ) );
        else if( aType.equals( sPivotTypeName ) )
            xHandler.set( new OoxPivotTableFragment( getWorksheetHelper(), aFragPath ) );

        if( xHandler.is() )
            getFilter()->importFragment( xHandler, aFragPath );
    }
}

void OoxWorksheetFragment::importColumn( const AttributeList& rAttribs )
{
    OoxColumnData aData;
    aData.mnFirstCol = rAttribs.getInteger( XML_min, -1 );
    aData.mnLastCol = rAttribs.getInteger( XML_max, -1 );
    aData.mfWidth = rAttribs.getDouble( XML_width, 0.0 );
    aData.mnXfId = rAttribs.getInteger( XML_style, -1 );
    aData.mnLevel = rAttribs.getInteger( XML_outlineLevel, 0 );
    aData.mbHidden = rAttribs.getBool( XML_hidden, false );
    aData.mbCollapsed = rAttribs.getBool( XML_collapsed, false );
    // set column properties in the current sheet
    setColumnData( aData );
}

void OoxWorksheetFragment::importMergeCell( const AttributeList& rAttribs )
{
    CellRangeAddress aRange;
    if( getAddressConverter().convertToCellRange( aRange, rAttribs.getString( XML_ref ), getSheetIndex(), true ) )
        setMergedRange( aRange );
}

void OoxWorksheetFragment::importSheetPr( const AttributeList& /*rAttribs*/ )
{
    // TODO - fill it.
}

void OoxWorksheetFragment::importOutlinePr( const AttributeList& rAttribs )
{
    setOutlineSummarySymbols(
        rAttribs.getBool( XML_summaryRight, true ),
        rAttribs.getBool( XML_summaryBelow, true ) );
}

void OoxWorksheetFragment::importPageSetUpPr( const AttributeList& rAttribs )
{
//    bool bAutoBreaks = rAttribs.getBool( XML_autoPageBreaks, true );
    getPageStyle().setFitToPagesMode( rAttribs.getBool( XML_fitToPage, false ) );
}

void OoxWorksheetFragment::importSheetFormatPr( const AttributeList& rAttribs )
{
    // default column settings
    setBaseColumnWidth( rAttribs.getInteger( XML_baseColWidth, 8 ) );
    setDefaultColumnWidth( rAttribs.getDouble( XML_defaultColWidth, 0.0 ) );
    // default row settings
    setDefaultRowSettings(
        rAttribs.getDouble( XML_defaultRowHeight, 0.0 ),
        rAttribs.getBool( XML_zeroHeight, false ) );
}

void OoxWorksheetFragment::importBrk( const AttributeList& rAttribs )
{
    OoxPageBreakData aData;
    aData.mnColRow = rAttribs.getInteger( XML_id, -1 );
    aData.mbManual = rAttribs.getBool( XML_man, false );
    switch( getPreviousContext() )
    {
        case XLS_TOKEN( colBreaks ):    convertPageBreak( aData, false );   break;
        case XLS_TOKEN( rowBreaks ):    convertPageBreak( aData, true );    break;
    }
}

void OoxWorksheetFragment::importHyperlink( const AttributeList& rAttribs )
{
    OoxHyperlinkData aData;
    if( getAddressConverter().convertToCellRange( aData.maRange, rAttribs.getString( XML_ref ), getSheetIndex(), true ) )
    {
        if( const Relation* pRel = getRelations()->getRelationById( rAttribs.getString( R_TOKEN( id ) ) ).get() )
            aData.maTarget = pRel->msTarget;
        aData.maLocation = rAttribs.getString( XML_location );
        aData.maDisplay = rAttribs.getString( XML_display );
        aData.maTooltip = rAttribs.getString( XML_tooltip );
        setHyperlink( aData );
    }
}

// ============================================================================

BiffWorksheetFragment::BiffWorksheetFragment( const GlobalDataHelper& rGlobalData,
        WorksheetType eSheetType, const Reference< XSpreadsheet >& rxSheet, sal_uInt32 nSheet ) :
    WorksheetHelper( rGlobalData, eSheetType, rxSheet, nSheet )
{
}

bool BiffWorksheetFragment::importFragment( BiffInputStream& rStrm )
{
    // create a SheetDataContext object that implements cell import
    BiffSheetDataContext aSheetData( getWorksheetHelper() );
    // create a SheetViewsContext object that implements view settings import
    BiffSheetViewsContext aSheetView( getWorksheetHelper() );

    ExternalLinkBuffer& rExtLinks = getExternalLinks();
    PageStyle& rPageStyle = getPageStyle();

    // process all record in this sheet fragment
    while( rStrm.startNextRecord() && (rStrm.getRecId() != BIFF_ID_EOF) )
    {
        sal_uInt16 nRecId = rStrm.getRecId();

        if( BiffHelper::isBofRecord( nRecId ) )
        {
            // skip unknown embedded fragments (BOF/EOF blocks)
            BiffHelper::skipFragment( rStrm );
        }
        else
        {
            // cache core stream position to detect if record is already processed
            sal_Int64 nStrmPos = rStrm.getCoreStreamPos();

            switch( nRecId )
            {
                // records in all BIFF versions
                case BIFF_ID_BOTTOMMARGIN:      rPageStyle.importBottomMargin( rStrm );     break;
                case BIFF_ID_DEFCOLWIDTH:       importDefColWidth( rStrm );                 break;
                case BIFF_ID_FOOTER:            rPageStyle.importFooter( rStrm );           break;
                case BIFF_ID_HEADER:            rPageStyle.importHeader( rStrm );           break;
                case BIFF_ID_HORPAGEBREAKS:     importPageBreaks( rStrm );                  break;
                case BIFF_ID_LEFTMARGIN:        rPageStyle.importLeftMargin( rStrm );       break;
                case BIFF_ID_PRINTGRIDLINES:    rPageStyle.importPrintGridLines( rStrm );   break;
                case BIFF_ID_PRINTHEADERS:      rPageStyle.importPrintHeaders( rStrm );     break;
                case BIFF_ID_RIGHTMARGIN:       rPageStyle.importRightMargin( rStrm );      break;
                case BIFF_ID_TOPMARGIN:         rPageStyle.importTopMargin( rStrm );        break;
                case BIFF_ID_VERPAGEBREAKS:     importPageBreaks( rStrm );                  break;

                // BIFF specific records
                default: switch( getBiff() )
                {
                    case BIFF2: switch( nRecId )
                    {
                        case BIFF_ID_COLUMNDEFAULT: importColumnDefault( rStrm );           break;
                        case BIFF_ID_COLWIDTH:      importColWidth( rStrm );                break;
                        case BIFF2_ID_DEFROWHEIGHT: importDefRowHeight( rStrm );            break;
                    }
                    break;

                    case BIFF3: switch( nRecId )
                    {
                        case BIFF_ID_COLINFO:       importColInfo( rStrm );                 break;
                        case BIFF_ID_DEFCOLWIDTH:   importDefColWidth( rStrm );             break;
                        case BIFF3_ID_DEFROWHEIGHT: importDefRowHeight( rStrm );            break;
                        case BIFF_ID_HCENTER:       rPageStyle.importHorCenter( rStrm );    break;
                        case BIFF_ID_VCENTER:       rPageStyle.importVerCenter( rStrm );    break;
                        case BIFF_ID_WSBOOL:        importWsBool( rStrm );                  break;
                    }
                    break;

                    case BIFF4: switch( nRecId )
                    {
                        case BIFF_ID_COLINFO:       importColInfo( rStrm );                 break;
                        case BIFF3_ID_DEFROWHEIGHT: importDefRowHeight( rStrm );            break;
                        case BIFF3_ID_EXTERNNAME:   rExtLinks.importExternName( rStrm );    break;
                        case BIFF_ID_EXTERNSHEET:   rExtLinks.importExternSheet( rStrm );   break;
                        case BIFF_ID_HCENTER:       rPageStyle.importHorCenter( rStrm );    break;
                        case BIFF_ID_SETUP:         rPageStyle.importSetup( rStrm );        break;
                        case BIFF_ID_STANDARDWIDTH: importStandardWidth( rStrm );           break;
                        case BIFF_ID_VCENTER:       rPageStyle.importVerCenter( rStrm );    break;
                        case BIFF_ID_WSBOOL:        importWsBool( rStrm );                  break;
                    }
                    break;

                    case BIFF5: switch( nRecId )
                    {
                        case BIFF_ID_COLINFO:       importColInfo( rStrm );                 break;
                        case BIFF3_ID_DEFROWHEIGHT: importDefRowHeight( rStrm );            break;
                        case BIFF5_ID_EXTERNNAME:   rExtLinks.importExternName( rStrm );    break;
                        case BIFF_ID_EXTERNSHEET:   rExtLinks.importExternSheet( rStrm );   break;
                        case BIFF_ID_HCENTER:       rPageStyle.importHorCenter( rStrm );    break;
                        case BIFF_ID_MERGEDCELLS:   importMergedCells( rStrm );             break;  // #i62300# also in BIFF5
                        case BIFF_ID_SETUP:         rPageStyle.importSetup( rStrm );        break;
                        case BIFF_ID_STANDARDWIDTH: importStandardWidth( rStrm );           break;
                        case BIFF_ID_VCENTER:       rPageStyle.importVerCenter( rStrm );    break;
                        case BIFF_ID_WSBOOL:        importWsBool( rStrm );                  break;
                    }
                    break;

                    case BIFF8: switch( nRecId )
                    {
                        case BIFF_ID_COLINFO:       importColInfo( rStrm );                 break;
                        case BIFF3_ID_DEFROWHEIGHT: importDefRowHeight( rStrm );            break;
                        case BIFF_ID_HCENTER:       rPageStyle.importHorCenter( rStrm );    break;
                        case BIFF_ID_HLINK:         importHlink( rStrm );                   break;
                        case BIFF_ID_MERGEDCELLS:   importMergedCells( rStrm );             break;
                        case BIFF_ID_SETUP:         rPageStyle.importSetup( rStrm );        break;
                        case BIFF_ID_STANDARDWIDTH: importStandardWidth( rStrm );           break;
                        case BIFF_ID_VCENTER:       rPageStyle.importVerCenter( rStrm );    break;
                        case BIFF_ID_WSBOOL:        importWsBool( rStrm );                  break;
                    }
                    break;

                    case BIFF_UNKNOWN: break;
                }
            }

            // record not processed, try cell records
            if( rStrm.getCoreStreamPos() == nStrmPos )
                aSheetData.importRecord( rStrm );
            // record not processed, try view settings records
            if( rStrm.getCoreStreamPos() == nStrmPos )
                aSheetView.importRecord( rStrm );
        }
    }

    // final processing (column/row settings, etc), and leave
    finalizeWorksheetImport();
    return rStrm.getRecId() == BIFF_ID_EOF;
}

void BiffWorksheetFragment::importColInfo( BiffInputStream& rStrm )
{
    sal_uInt16 nFirstCol, nLastCol, nWidth, nXfId, nFlags;
    rStrm >> nFirstCol >> nLastCol >> nWidth >> nXfId >> nFlags;

    OoxColumnData aData;
    // column indexes are 0-based in BIFF, but OoxColumnData expects 1-based
    aData.mnFirstCol = static_cast< sal_Int32 >( nFirstCol ) + 1;
    aData.mnLastCol = static_cast< sal_Int32 >( nLastCol ) + 1;
    // width is stored as 1/256th of a character in BIFF, convert to entire character
    aData.mfWidth = static_cast< double >( nWidth ) / 256.0;
    aData.mnXfId = nXfId;
    aData.mnLevel = extractValue< sal_Int32 >( nFlags, 8, 3 );
    aData.mbHidden = getFlag( nFlags, BIFF_COLINFO_HIDDEN );
    aData.mbCollapsed = getFlag( nFlags, BIFF_COLINFO_COLLAPSED );
    // set column properties in the current sheet
    setColumnData( aData );
}

void BiffWorksheetFragment::importColumnDefault( BiffInputStream& rStrm )
{
    sal_uInt16 nFirstCol, nLastCol, nXfId;
    rStrm >> nFirstCol >> nLastCol >> nXfId;
    convertColumnFormat( nFirstCol, nLastCol, nXfId );
}

void BiffWorksheetFragment::importColWidth( BiffInputStream& rStrm )
{
    sal_uInt8 nFirstCol, nLastCol;
    sal_uInt16 nWidth;
    rStrm >> nFirstCol >> nLastCol >> nWidth;

    OoxColumnData aData;
    // column indexes are 0-based in BIFF, but OoxColumnData expects 1-based
    aData.mnFirstCol = static_cast< sal_Int32 >( nFirstCol ) + 1;
    aData.mnLastCol = static_cast< sal_Int32 >( nLastCol ) + 1;
    // width is stored as 1/256th of a character in BIFF, convert to entire character
    aData.mfWidth = static_cast< double >( nWidth ) / 256.0;
    // set column properties in the current sheet
    setColumnData( aData );
}

void BiffWorksheetFragment::importDefColWidth( BiffInputStream& rStrm )
{
    /*  Stored as entire number of characters without padding pixels, which
        will be added in setBaseColumnWidth(). Call has no effect, if a
        width has already been set from the STANDARDWIDTH record. */
    setBaseColumnWidth( rStrm.readuInt16() );
}

void BiffWorksheetFragment::importDefRowHeight( BiffInputStream& rStrm )
{
    sal_uInt16 nFlags = BIFF_DEFROW_UNSYNCED, nHeight;
    if( getBiff() != BIFF2 )
        rStrm >> nFlags;
    rStrm >> nHeight;
    if( getBiff() == BIFF2 )
        nHeight &= BIFF2_DEFROW_MASK;
    // row height is in twips in BIFF, convert to points
    setDefaultRowSettings( nHeight / 20.0, getFlag( nFlags, BIFF_DEFROW_HIDDEN ) );
}

void BiffWorksheetFragment::importHlink( BiffInputStream& rStrm )
{
    OoxHyperlinkData aData;

    // read cell range for the hyperlink
    BiffRange aBiffRange;
    rStrm >> aBiffRange;
    // #i80006# Excel silently ignores invalid hi-byte of column index (TODO: everywhere?)
    aBiffRange.maFirst.mnCol &= 0xFF;
    aBiffRange.maLast.mnCol &= 0xFF;
    if( !getAddressConverter().convertToCellRange( aData.maRange, aBiffRange, getSheetIndex(), true ) )
        return;

    BiffGuid aGuid;
    sal_uInt32 nId, nFlags;
    rStrm >> aGuid >> nId >> nFlags;

    OSL_ENSURE( aGuid == BiffHelper::maGuidStdHlink, "BiffWorksheetFragment::importHlink - unexpected header GUID" );
    OSL_ENSURE( nId == 2, "BiffWorksheetFragment::importHlink - unexpected header identifier" );
    if( !(aGuid == BiffHelper::maGuidStdHlink) )
        return;

    // display string
    if( getFlag( nFlags, BIFF_HLINK_DISPLAY ) )
        aData.maDisplay = readHlinkString( rStrm, true );
    // target frame (ignore) !TODO: DISPLAY/FRAME - right order? (never seen them together)
    if( getFlag( nFlags, BIFF_HLINK_FRAME ) )
        ignoreHlinkString( rStrm, true );

    // target
    if( getFlag( nFlags, BIFF_HLINK_TARGET ) )
    {
        if( getFlag( nFlags, BIFF_HLINK_UNC ) )
        {
            // UNC path
            OSL_ENSURE( getFlag( nFlags, BIFF_HLINK_ABS ), "BiffWorksheetFragment::importHlink - UNC link not absolute" );
            aData.maTarget = readHlinkString( rStrm, true );
        }
        else
        {
            rStrm >> aGuid;
            if( aGuid == BiffHelper::maGuidFileMoniker )
            {
                // file name, maybe relative and with directory up-count
                sal_Int16 nUpLevels;
                rStrm >> nUpLevels;
                OSL_ENSURE( (nUpLevels == 0) || !getFlag( nFlags, BIFF_HLINK_ABS ), "BiffWorksheetFragment::importHlink - absolute filename with upcount" );
                OUString aShortName = readHlinkString( rStrm, false );
                rStrm.ignore( 24 );
                if( rStrm.readInt32() > 0 )
                {
                    sal_Int32 nStrLen = rStrm.readInt32() / 2;  // byte count to char count
                    rStrm.ignore( 2 );
                    aData.maTarget = readHlinkString( rStrm, nStrLen, true );
                }
                if( aData.maTarget.getLength() == 0 )
                    aData.maTarget = aShortName;
                if( !getFlag( nFlags, BIFF_HLINK_ABS ) )
                    for( sal_uInt16 nLevel = 0; nLevel < nUpLevels; ++nLevel )
                        aData.maTarget = CREATE_OUSTRING( "..\\" ) + aData.maTarget;
            }
            else if( aGuid == BiffHelper::maGuidUrlMoniker )
            {
                // URL, maybe relative and with leading '../'
                sal_Int32 nStrLen = rStrm.readInt32() / 2;  // byte count to char count
                aData.maTarget = readHlinkString( rStrm, nStrLen, true );
            }
            else
            {
                OSL_ENSURE( false, "BiffWorksheetFragment::importHlink - unknown content GUID" );
                return;
            }
        }
    }

    // target location
    if( getFlag( nFlags, BIFF_HLINK_LOC ) )
        aData.maLocation = readHlinkString( rStrm, true );

    OSL_ENSURE( rStrm.getRecLeft() == 0, "BiffWorksheetFragment::importHlink - unknown record data" );

    // try to read the SCREENTIP record
    if( (rStrm.getNextRecId() == BIFF_ID_SCREENTIP) && rStrm.startNextRecord() )
    {
        rStrm.ignore( 2 );      // repeated record id
        // the cell range, again
        rStrm >> aBiffRange;
        CellRangeAddress aRange;
        if( getAddressConverter().convertToCellRange( aRange, aBiffRange, getSheetIndex(), true ) &&
            (aRange.StartColumn == aData.maRange.StartColumn) &&
            (aRange.StartRow == aData.maRange.StartRow) &&
            (aRange.EndColumn == aData.maRange.EndColumn) &&
            (aRange.EndRow == aData.maRange.EndRow) )
        {
            /*  This time, we have no string length, no flag field, and a
                null-terminated 16-bit character array. */
            aData.maTooltip = rStrm.readUnicodeArray( static_cast< sal_uInt16 >( rStrm.getRecLeft() / 2 ) );
        }
    }

    // store the hyperlink settings
    setHyperlink( aData );
}

void BiffWorksheetFragment::importMergedCells( BiffInputStream& rStrm )
{
    BiffRangeList aBiffRanges;
    rStrm >> aBiffRanges;
    ::std::vector< CellRangeAddress > aRanges;
    getAddressConverter().convertToCellRangeList( aRanges, aBiffRanges, getSheetIndex(), true );
    for( ::std::vector< CellRangeAddress >::const_iterator aIt = aRanges.begin(), aEnd = aRanges.end(); aIt != aEnd; ++aIt )
        setMergedRange( *aIt );
}

void BiffWorksheetFragment::importPageBreaks( BiffInputStream& rStrm )
{
    bool bRowBreak = false;
    switch( rStrm.getRecId() )
    {
        case BIFF_ID_HORPAGEBREAKS: bRowBreak = true;   break;
        case BIFF_ID_VERPAGEBREAKS: bRowBreak = false;  break;
        default:
            OSL_ENSURE( false, "BiffWorksheetFragment::importPageBreaks - unknown record" );
            return;
    }

    OoxPageBreakData aData;
    aData.mbManual = true;              // only manual breaks stored in BIFF
    bool bBiff8 = getBiff() == BIFF8;   // ignore start/end columns or rows in BIFF8

    sal_uInt16 nCount;
    rStrm >> nCount;
    for( sal_uInt16 nIndex = 0; rStrm.isValid() && (nIndex < nCount); ++nIndex )
    {
        aData.mnColRow = rStrm.readuInt16();
        convertPageBreak( aData, bRowBreak );
        if( bBiff8 )
            rStrm.ignore( 4 );
    }
}

void BiffWorksheetFragment::importStandardWidth( BiffInputStream& rStrm )
{
    sal_uInt16 nWidth;
    rStrm >> nWidth;
    // width is stored as 1/256th of a character in BIFF, convert to entire character
    double fWidth = static_cast< double >( nWidth ) / 256.0;
    // set as default width, will override the width from DEFCOLWIDTH record
    setDefaultColumnWidth( fWidth );
}

void BiffWorksheetFragment::importWsBool( BiffInputStream& rStrm )
{
    sal_uInt16 nFlags;
    rStrm >> nFlags;
    // position of outline summary symbols
    setOutlineSummarySymbols(
        getFlag( nFlags, BIFF_WSBOOL_COLRIGHT ),
        getFlag( nFlags, BIFF_WSBOOL_ROWBELOW ) );
    // fit printout to width/height
    getPageStyle().setFitToPagesMode( getFlag( nFlags, BIFF_WSBOOL_FITTOPAGES ) );
}

OUString BiffWorksheetFragment::readHlinkString( BiffInputStream& rStrm, sal_Int32 nChars, bool bUnicode )
{
    OUString aRet;
    if( nChars > 0 )
    {
        sal_uInt16 nReadChars = getLimitedValue< sal_uInt16, sal_Int32 >( nChars, 0, SAL_MAX_UINT16 );
        aRet = bUnicode ?
            rStrm.readUnicodeArray( nReadChars ) :
            rStrm.readCharArray( nReadChars, getTextEncoding() );
        // ignore remaining chars
        sal_uInt32 nIgnore = static_cast< sal_uInt32 >( nChars - nReadChars );
        rStrm.ignore( bUnicode ? (nIgnore * 2) : nIgnore );
    }
    return aRet;
}

OUString BiffWorksheetFragment::readHlinkString( BiffInputStream& rStrm, bool bUnicode )
{
    return readHlinkString( rStrm, rStrm.readInt32(), bUnicode );
}

void BiffWorksheetFragment::ignoreHlinkString( BiffInputStream& rStrm, sal_Int32 nChars, bool bUnicode )
{
    if( nChars > 0 )
        rStrm.ignore( static_cast< sal_uInt32 >( bUnicode ? (nChars * 2) : nChars ) );
}

void BiffWorksheetFragment::ignoreHlinkString( BiffInputStream& rStrm, bool bUnicode )
{
    ignoreHlinkString( rStrm, rStrm.readInt32(), bUnicode );
}

// ============================================================================

} // namespace xls
} // namespace oox

