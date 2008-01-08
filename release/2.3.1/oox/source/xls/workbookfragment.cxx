/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: workbookfragment.cxx,v $
 *
 *  $Revision: 1.1.2.48 $
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

#include "oox/xls/workbookfragment.hxx"
#include <com/sun/star/table/CellAddress.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include "oox/core/propertyset.hxx"
#include "oox/drawingml/themefragmenthandler.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/condformatbuffer.hxx"
#include "oox/xls/connectionsfragment.hxx"
#include "oox/xls/externallinkbuffer.hxx"
#include "oox/xls/externallinkfragment.hxx"
#include "oox/xls/pivotcachefragment.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/sharedstringsfragment.hxx"
#include "oox/xls/stylesfragment.hxx"
#include "oox/xls/themebuffer.hxx"
#include "oox/xls/viewsettings.hxx"
#include "oox/xls/worksheetbuffer.hxx"
#include "oox/xls/worksheetfragment.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::sheet::XSpreadsheetDocument;
using ::com::sun::star::table::CellAddress;
using ::oox::core::AttributeList;
using ::oox::core::PropertySet;
using ::oox::drawingml::ThemeFragmentHandler;

namespace oox {
namespace xls {

// ============================================================================

OoxWorkbookFragment::OoxWorkbookFragment(
        const GlobalDataHelper& rGlobalData, const OUString& rFragmentPath ) :
    GlobalFragmentBase( rGlobalData, rFragmentPath )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxWorkbookFragment::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XML_ROOT_CONTEXT:
            return  (nElement == XLS_TOKEN( workbook ));
        case XLS_TOKEN( workbook ):
            return  (nElement == XLS_TOKEN( sheets )) ||
                    (nElement == XLS_TOKEN( bookViews )) ||
                    (nElement == XLS_TOKEN( externalReferences )) ||
                    (nElement == XLS_TOKEN( definedNames )) ||
                    (nElement == XLS_TOKEN( pivotCaches ));
        case XLS_TOKEN( sheets ):
            return  (nElement == XLS_TOKEN( sheet ));
        case XLS_TOKEN( bookViews ):
            return  (nElement == XLS_TOKEN( workbookView ));
        case XLS_TOKEN( externalReferences ):
            return  (nElement == XLS_TOKEN( externalReference ));
        case XLS_TOKEN( definedNames ):
            return  (nElement == XLS_TOKEN( definedName ));
        case XLS_TOKEN( pivotCaches ):
            return  (nElement == XLS_TOKEN( pivotCache ));
    }
    return false;
}

void OoxWorkbookFragment::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sheet ):
            getWorksheets().importSheet( rAttribs );
        break;
        case XLS_TOKEN( workbookView ):
            importWorkbookView( rAttribs );
        break;
        case XLS_TOKEN( externalReference ):
            importExternalReference( rAttribs );
        break;
        case XLS_TOKEN( definedName ):
            mxCurrName = getDefinedNames().importDefinedName( rAttribs );
        break;
        case XLS_TOKEN( pivotCache ):
            importPivotCache( rAttribs );
        break;
    }
}

void OoxWorkbookFragment::onEndElement( const OUString& rChars )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( definedName ):
            if( mxCurrName.get() ) mxCurrName->setFormula( rChars );
        break;
        case XLS_TOKEN( workbook ):
            finalizeImport();
        break;
    }
}

// private --------------------------------------------------------------------

void OoxWorkbookFragment::importWorkbookView( const AttributeList& rAttribs )
{
    OoxWorkbookViewData& rData = getViewSettings().createWorkbookViewData();
    rData.mnWinX = rAttribs.getInteger( XML_xWindow, 0 );
    rData.mnWinY = rAttribs.getInteger( XML_yWindow, 0 );
    rData.mnWinWidth = rAttribs.getInteger( XML_windowWidth, 0 );
    rData.mnWinHeight = rAttribs.getInteger( XML_windowHeight, 0 );
    rData.mnActiveSheet = rAttribs.getInteger( XML_activeTab, 0 );
    rData.mnFirstVisSheet = rAttribs.getInteger( XML_firstSheet, 0 );
    rData.mnTabBarWidth = rAttribs.getInteger( XML_tabRatio, 600 );
    rData.mnVisibility = rAttribs.getToken( XML_visibility, XML_visible );
    rData.mbShowTabBar = rAttribs.getBool( XML_showSheetTabs, true );
    rData.mbShowHorScroll = rAttribs.getBool( XML_showHorizontalScroll, true );
    rData.mbShowVerScroll = rAttribs.getBool( XML_showVerticalScroll, true );
    rData.mbMinimized = rAttribs.getBool( XML_minimized, false );
}

void OoxWorkbookFragment::importExternalReference( const AttributeList& rAttribs )
{
    ExternalLinkRef xExtLink = getExternalLinks().importExternalReference( rAttribs );
    OUString aFragmentPath = getFragmentPathByRelId( rAttribs.getString( R_TOKEN( id ) ) );
    if( xExtLink.get() && (aFragmentPath.getLength() > 0) )
    {
        Reference< XFastDocumentHandler > xHandler( new OoxExternalLinkFragment( getGlobalData(), *xExtLink, aFragmentPath ) );
        getFilter()->importFragment( xHandler, aFragmentPath );
    }
}

void OoxWorkbookFragment::importPivotCache( const AttributeList& rAttribs )
{
    OUString aFragPath = getFragmentPathByRelId( rAttribs.getString( R_TOKEN( id ) ) );
    if( (aFragPath.getLength() > 0) && rAttribs.hasAttribute( XML_cacheId ) )
    {
        sal_uInt32 nCacheId = rAttribs.getUnsignedInteger( XML_cacheId, 0 );
        Reference< XFastDocumentHandler > xHandler = new OoxPivotCacheFragment( getGlobalData(), aFragPath, nCacheId );
        getFilter()->importFragment( xHandler, aFragPath );
    }
}

void OoxWorkbookFragment::finalizeImport()
{
    OUString aFragmentPath;

    // read the theme substream
    aFragmentPath = getFragmentPathByType( CREATE_RELATIONS_TYPE( "theme" ) );
    if( aFragmentPath.getLength() > 0 )
    {
        Reference< XFastDocumentHandler > xHandler = new ThemeFragmentHandler(
            getFilter(), aFragmentPath, getTheme().getCoreTheme() );
        getFilter()->importFragment( xHandler, aFragmentPath );
    }

    // read the styles substream (requires finalized theme buffer)
    aFragmentPath = getFragmentPathByType( CREATE_RELATIONS_TYPE( "styles" ) );
    if( aFragmentPath.getLength() > 0 )
    {
        Reference< XFastDocumentHandler > xHandler = new OoxStylesFragment( getGlobalData(), aFragmentPath );
        getFilter()->importFragment( xHandler, aFragmentPath );
    }

    // read the shared string table substream (requires finalized styles buffer)
    aFragmentPath = getFragmentPathByType( CREATE_RELATIONS_TYPE( "sharedStrings" ) );
    if( aFragmentPath.getLength() > 0 )
    {
        Reference< XFastDocumentHandler > xHandler = new OoxSharedStringsFragment( getGlobalData(), aFragmentPath );
        getFilter()->importFragment( xHandler, aFragmentPath );
    }

    // read the connections substream
    aFragmentPath = getFragmentPathByType( CREATE_RELATIONS_TYPE( "connections" ) );
    if( aFragmentPath.getLength() > 0 )
    {
        Reference< XFastDocumentHandler > xHandler = new OoxConnectionsFragment( getGlobalData(), aFragmentPath );
        getFilter()->importFragment( xHandler, aFragmentPath );
    }

    // create all defined names
    getDefinedNames().finalizeImport();

    // load all worksheets
    WorksheetBuffer& rWorksheets = getWorksheets();
    for( sal_Int32 nSheet = 0, nCount = rWorksheets.getInternalSheetCount(); nSheet < nCount; ++nSheet )
    {
        // get reference to the current worksheet in the document model
        Reference< XSpreadsheet > xSheet = getSheet( nSheet );
        OSL_ENSURE( xSheet.is(), "OoxWorkbookFragment::finalizeImport - cannot access sheet in document model" );

        // get fragment path of the sheet
        aFragmentPath = getFragmentPathByRelId( rWorksheets.getSheetRelId( nSheet ) );
        OSL_ENSURE( aFragmentPath.getLength() > 0, "OoxWorkbookFragment::finalizeImport - cannot access sheet fragment" );

        // load the worksheet substream
        if( xSheet.is() && (aFragmentPath.getLength() > 0) )
        {
            Reference< XFastDocumentHandler > xHandler =
                new OoxWorksheetFragment( getGlobalData(), aFragmentPath, xSheet, nSheet );
            getFilter()->importFragment( xHandler, aFragmentPath );
        }
    }

    // document and sheet view settings, after loading all sheets
    getViewSettings().finalizeImport();

    getPivotTables().finalizeImport();
    getCondFormats().finalizeImport();
}

// ============================================================================

namespace {

BiffDecoderRef lclImportFilePass_XOR( const GlobalDataHelper& rGlobalData, BiffInputStream& rStrm )
{
    BiffDecoderRef xDecoder;
    OSL_ENSURE( rStrm.getRecLeft() == 4, "lclImportFilePass_XOR - wrong record size" );
    if( rStrm.getRecLeft() == 4 )
    {
        sal_uInt16 nBaseKey, nHash;
        rStrm >> nBaseKey >> nHash;
        xDecoder.reset( new BiffDecoder_XOR( rGlobalData, nBaseKey, nHash ) );
    }
    return xDecoder;
}

BiffDecoderRef lclImportFilePass_RCF( const GlobalDataHelper& rGlobalData, BiffInputStream& rStrm )
{
    BiffDecoderRef xDecoder;
    OSL_ENSURE( rStrm.getRecLeft() == 48, "lclImportFilePass_RCF - wrong record size" );
    if( rStrm.getRecLeft() == 48 )
    {
        sal_uInt8 pnDocId[ 16 ];
        sal_uInt8 pnSaltData[ 16 ];
        sal_uInt8 pnSaltHash[ 16 ];
        rStrm.read( pnDocId, 16 );
        rStrm.read( pnSaltData, 16 );
        rStrm.read( pnSaltHash, 16 );
        xDecoder.reset( new BiffDecoder_RCF( rGlobalData, pnDocId, pnSaltData, pnSaltHash ) );
    }
    return xDecoder;
}

BiffDecoderRef lclImportFilePass_Strong( const GlobalDataHelper& /*rGlobalData*/, BiffInputStream& /*rStrm*/ )
{
    // not supported
    return BiffDecoderRef();
}

BiffDecoderRef lclImportFilePass2( const GlobalDataHelper& rGlobalData, BiffInputStream& rStrm )
{
    return lclImportFilePass_XOR( rGlobalData, rStrm );
}

BiffDecoderRef lclImportFilePass8( const GlobalDataHelper& rGlobalData, BiffInputStream& rStrm )
{
    BiffDecoderRef xDecoder;

    switch( rStrm.readuInt16() )
    {
        case BIFF_FILEPASS_BIFF2:
            xDecoder = lclImportFilePass_XOR( rGlobalData, rStrm );
        break;

        case BIFF_FILEPASS_BIFF8:
        {
            rStrm.ignore( 2 );
            switch( rStrm.readuInt16() )
            {
                case BIFF_FILEPASS_BIFF8_RCF:
                    xDecoder = lclImportFilePass_RCF( rGlobalData, rStrm );
                break;
                case BIFF_FILEPASS_BIFF8_STRONG:
                    xDecoder = lclImportFilePass_Strong( rGlobalData, rStrm );
                break;
                default:
                    OSL_ENSURE( false, "lclImportFilePass8 - unknown BIFF8 encryption sub mode" );
            }
        }
        break;

        default:
            OSL_ENSURE( false, "lclImportFilePass8 - unknown encryption mode" );
    }

    return xDecoder;
}

} // namespace

// ----------------------------------------------------------------------------

BiffWorkbookFragment::BiffWorkbookFragment( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

bool BiffWorkbookFragment::importWorkbook( BiffInputStream& rStrm )
{
    bool bRet = false;

    BiffFragmentType eFragment = BiffHelper::startFragment( rStrm, getBiff() );
    switch( eFragment )
    {
        case BIFF_FRAGMENT_GLOBALS:
        {
            // import workbook globals fragment and create sheets in document
            bRet = importGlobalsFragment( rStrm );
            WorksheetBuffer& rWorksheets = getWorksheets();
            // load sheet fragments (do not return false in bRet on missing/broken sheets)
            bool bNextSheet = bRet;
            for( sal_Int32 nSheet = 0, nCount = rWorksheets.getInternalSheetCount(); bNextSheet && (nSheet < nCount); ++nSheet )
            {
                // try to start a new sheet fragment
                BiffFragmentType eSheetFragment = BiffHelper::startFragment( rStrm, getBiff() );
                bNextSheet = importSheetFragment( rStrm, eSheetFragment, nSheet );
            }
        }
        break;

        case BIFF_FRAGMENT_WORKSPACE:
        {
            bRet = importWorkspaceFragment( rStrm );
            // sheets are embedded in workspace fragment, nothing to do here
        }
        break;

        case BIFF_FRAGMENT_WORKSHEET:
        case BIFF_FRAGMENT_CHART:
        case BIFF_FRAGMENT_MACRO:
        {
            /*  Single sheet without globals
                - #i62752# possible in all BIFF versions
                - do not return false in bRet on missing/broken sheets. */
            importSheetFragment( rStrm, eFragment, 0 );
            // success, even if stream is broken
            bRet = true;
        }
        break;

        default:;
    }

    // document and sheet view settings, collected for all sheets
    getViewSettings().finalizeImport();

    return bRet;
}

bool BiffWorkbookFragment::importGlobalsFragment( BiffInputStream& rStrm )
{
    WorksheetBuffer& rWorksheets = getWorksheets();
    SharedStringsBuffer& rSharedStrings = getSharedStrings();
    StylesBuffer& rStyles = getStyles();
    ExternalLinkBuffer& rExtLinks = getExternalLinks();
    DefinedNamesBuffer& rDefNames = getDefinedNames();

    bool bRet = true;
    bool bLoop = true;
    while( bRet && bLoop && rStrm.startNextRecord() && (rStrm.getRecId() != BIFF_ID_EOF) )
    {
        sal_uInt16 nRecId = rStrm.getRecId();

        /*  #i56376# BIFF5-BIFF8: If an EOF record for globals is missing,
            simulate it. The issue is about a document where the sheet fragment
            starts directly after the EXTSST record, without terminating the
            globals fragment with an EOF record. */
        if( BiffHelper::isBofRecord( nRecId ) ) switch( getBiff() )
        {
            case BIFF2:
            case BIFF3:
            case BIFF4:
                // BIFF2-BIFF4: skip unknown embedded fragments
                BiffHelper::skipFragment( rStrm );
            break;
            case BIFF5:
            case BIFF8:
                rStrm.rewindRecord();
                bLoop = false;
            break;
            case BIFF_UNKNOWN: break;
        }
        else switch( nRecId )
        {
            // records in all BIFF versions
            case BIFF_ID_CODEPAGE:      setCodePage( rStrm.readuInt16() );      break;
            case BIFF_ID_FILEPASS:      bRet = importFilePass( rStrm );         break;
            case BIFF_ID_WINDOW1:       importWindow1( rStrm );                 break;

            // BIFF specific records
            default: switch( getBiff() )
            {
                case BIFF2: switch( nRecId )
                {
                    case BIFF2_ID_FONT:         rStyles.importFont( rStrm );            break;
                    case BIFF_ID_FONTCOLOR:     rStyles.importFontColor( rStrm );       break;
                    case BIFF2_ID_FORMAT:       rStyles.importFormat( rStrm );          break;
                    case BIFF2_ID_NAME:         rDefNames.importName( rStrm );          break;
                    case BIFF2_ID_XF:           rStyles.importXf( rStrm );              break;
                }
                break;

                case BIFF3: switch( nRecId )
                {
                    case BIFF3_ID_FONT:         rStyles.importFont( rStrm );            break;
                    case BIFF2_ID_FORMAT:       rStyles.importFormat( rStrm );          break;
                    case BIFF3_ID_NAME:         rDefNames.importName( rStrm );          break;
                    case BIFF_ID_PALETTE:       rStyles.importPalette( rStrm );         break;
                    case BIFF_ID_STYLE:         rStyles.importStyle( rStrm );           break;
                    case BIFF3_ID_XF:           rStyles.importXf( rStrm );              break;
                }
                break;

                case BIFF4: switch( nRecId )
                {
                    case BIFF3_ID_FONT:         rStyles.importFont( rStrm );            break;
                    case BIFF4_ID_FORMAT:       rStyles.importFormat( rStrm );          break;
                    case BIFF3_ID_NAME:         rDefNames.importName( rStrm );          break;
                    case BIFF_ID_PALETTE:       rStyles.importPalette( rStrm );         break;
                    case BIFF_ID_STYLE:         rStyles.importStyle( rStrm );           break;
                    case BIFF4_ID_XF:           rStyles.importXf( rStrm );              break;
                }
                break;

                case BIFF5: switch( nRecId )
                {
                    case BIFF_ID_BOUNDSHEET:    rWorksheets.importBoundSheet( rStrm );  break;
                    case BIFF5_ID_EXTERNNAME:   rExtLinks.importExternName( rStrm );    break;
                    case BIFF_ID_EXTERNSHEET:   rExtLinks.importExternSheet( rStrm );   break;
                    case BIFF5_ID_FONT:         rStyles.importFont( rStrm );            break;
                    case BIFF4_ID_FORMAT:       rStyles.importFormat( rStrm );          break;
                    case BIFF5_ID_NAME:         rDefNames.importName( rStrm );          break;
                    case BIFF_ID_PALETTE:       rStyles.importPalette( rStrm );         break;
                    case BIFF_ID_STYLE:         rStyles.importStyle( rStrm );           break;
                    case BIFF5_ID_XF:           rStyles.importXf( rStrm );              break;
                }
                break;

                case BIFF8: switch( nRecId )
                {
                    case BIFF_ID_BOUNDSHEET:    rWorksheets.importBoundSheet( rStrm );  break;
                    case BIFF5_ID_EXTERNNAME:   rExtLinks.importExternName( rStrm );    break;
                    case BIFF_ID_EXTERNSHEET:   rExtLinks.importExternSheet( rStrm );   break;
                    case BIFF5_ID_FONT:         rStyles.importFont( rStrm );            break;
                    case BIFF4_ID_FORMAT:       rStyles.importFormat( rStrm );          break;
                    case BIFF5_ID_NAME:         rDefNames.importName( rStrm );          break;
                    case BIFF_ID_PALETTE:       rStyles.importPalette( rStrm );         break;
                    case BIFF_ID_SST:           rSharedStrings.importSst( rStrm );      break;
                    case BIFF_ID_STYLE:         rStyles.importStyle( rStrm );           break;
                    case BIFF_ID_SUPBOOK:       rExtLinks.importSupBook( rStrm );       break;
                    case BIFF5_ID_XF:           rStyles.importXf( rStrm );              break;
                }
                break;

                case BIFF_UNKNOWN: break;
            }
        }
    }

    // finalize global buffers
    rSharedStrings.finalizeImport();
    rStyles.finalizeImport();
    rDefNames.finalizeImport();
    return bRet;
}

bool BiffWorkbookFragment::importWorkspaceFragment( BiffInputStream& rStrm )
{
    // enable workbook mode, has not been set yet in BIFF4 workspace files
    setIsWorkbookFile();

    WorksheetBuffer& rWorksheets = getWorksheets();
    bool bRet = true;

    // import the workspace globals
    bool bLoop = true;
    while( bRet && bLoop && rStrm.startNextRecord() && (rStrm.getRecId() != BIFF_ID_EOF) )
    {
        switch( rStrm.getRecId() )
        {
            case BIFF_ID_BOUNDSHEET:    rWorksheets.importBoundSheet( rStrm );  break;
            case BIFF_ID_CODEPAGE:      setCodePage( rStrm.readuInt16() );      break;
            case BIFF_ID_FILEPASS:      bRet = importFilePass( rStrm );         break;
            case BIFF_ID_SHEETHEADER:   rStrm.rewindRecord(); bLoop = false;    break;
        }
    }

    // load sheet fragments (do not return false in bRet on missing/broken sheets)
    bool bNextSheet = bRet;
    for( sal_Int32 nSheet = 0, nCount = rWorksheets.getInternalSheetCount(); bNextSheet && (nSheet < nCount); ++nSheet )
    {
        // try to start a new sheet fragment (with leading SHEETHEADER record)
        bNextSheet = rStrm.startNextRecord() && (rStrm.getRecId() == BIFF_ID_SHEETHEADER);
        if( bNextSheet )
        {
            /*  Import SHEETHEADER record, read current sheet name (sheet substreams
                may not be in the same order as BOUNDSHEET records are). */
            rStrm.ignore( 4 );
            OUString aSheetName = rStrm.readByteString( false, getTextEncoding() );
            sal_Int32 nCurrSheet = rWorksheets.getFinalSheetIndex( aSheetName );
            // load the sheet fragment records
            BiffFragmentType eSheetFragment = BiffHelper::startFragment( rStrm, getBiff() );
            bNextSheet = importSheetFragment( rStrm, eSheetFragment, nCurrSheet );
            // do not return false in bRet on missing/broken sheets
        }
    }

    return bRet;
}

bool BiffWorkbookFragment::importSheetFragment( BiffInputStream& rStrm, BiffFragmentType eFragment, sal_Int32 nSheet )
{
    // #i11183# clear buffers that may be used per-sheet, e.g. in BIFF4 workspace files
    createBuffersPerSheet();

    // load the workbook globals fragment records in BIFF2-BIFF4
    if( getBiff() <= BIFF4 )
    {
        // set sheet index in defined names buffer to handle built-in names correctly
        getDefinedNames().setLocalSheetIndex( nSheet );
        // remember current record to seek back below
        sal_Int64 nRecHandle = rStrm.getRecHandle();
        importGlobalsFragment( rStrm );
        // rewind stream to fragment BOF record
        rStrm.startRecordByHandle( nRecHandle );
    }

    // get XSpreadsheet interface, force to skip sheet substream on error
    Reference< XSpreadsheet > xSheet = getSheet( nSheet );
    if( !xSheet.is() )
        eFragment = BIFF_FRAGMENT_EMPTYSHEET;

    // load the sheet fragment records
    bool bRet = false;
    switch( eFragment )
    {
        case BIFF_FRAGMENT_WORKSHEET:
            bRet = importWorksheetFragment( rStrm, xSheet, nSheet );
        break;
        case BIFF_FRAGMENT_CHART:
            bRet = importChartsheetFragment( rStrm, xSheet, nSheet );
        break;
        case BIFF_FRAGMENT_MACRO:
            bRet = importMacrosheetFragment( rStrm, xSheet, nSheet );
        break;
        case BIFF_FRAGMENT_EMPTYSHEET:
            bRet = BiffHelper::skipFragment( rStrm );
        break;
        default:;
    }
    return bRet;
}

bool BiffWorkbookFragment::importWorksheetFragment( BiffInputStream& rStrm, const Reference< XSpreadsheet >& rxSheet, sal_Int32 nSheet )
{
    BiffWorksheetFragment aFragment( getGlobalData(), SHEETTYPE_WORKSHEET, rxSheet, nSheet );
    return aFragment.importFragment( rStrm );
}

bool BiffWorkbookFragment::importChartsheetFragment( BiffInputStream& rStrm, const Reference< XSpreadsheet >& /*rxSheet*/, sal_Int32 /*nSheet*/ )
{
    return BiffHelper::skipFragment( rStrm );
}

bool BiffWorkbookFragment::importMacrosheetFragment( BiffInputStream& rStrm, const Reference< XSpreadsheet >& rxSheet, sal_Int32 nSheet )
{
    BiffWorksheetFragment aFragment( getGlobalData(), SHEETTYPE_MACRO, rxSheet, nSheet );
    return aFragment.importFragment( rStrm );
}

bool BiffWorkbookFragment::importFilePass( BiffInputStream& rStrm )
{
    rStrm.enableDecoder( false );
    BiffDecoderRef xDecoder = (getBiff() == BIFF8) ?
        lclImportFilePass8( getGlobalData(), rStrm ) : lclImportFilePass2( getGlobalData(), rStrm );

    // set decoder at import stream
    rStrm.setDecoder( xDecoder );
    //! TODO remember encryption state for export
//    rStrm.GetRoot().GetExtDocOptions().GetDocSettings().mbEncrypted = true;

    return xDecoder.get() && xDecoder->isValid();
}

void BiffWorkbookFragment::importWindow1( BiffInputStream& rStrm )
{
    sal_uInt16 nWinX, nWinY, nWinWidth, nWinHeight;
    rStrm >> nWinX >> nWinY >> nWinWidth >> nWinHeight;

    OoxWorkbookViewData& rData = getViewSettings().createWorkbookViewData();
    rData.mnWinX = nWinX;
    rData.mnWinY = nWinY;
    rData.mnWinWidth = nWinWidth;
    rData.mnWinHeight = nWinHeight;

    if( getBiff() <= BIFF4 )
    {
        sal_uInt8 nHidden;
        rStrm >> nHidden;
        rData.mnVisibility = (nHidden == 0) ? XML_visible : XML_hidden;
    }
    else
    {
        sal_uInt16 nFlags, nActiveTab, nFirstVisTab, nSelectCnt, nTabBarWidth;
        rStrm >> nFlags >> nActiveTab >> nFirstVisTab >> nSelectCnt >> nTabBarWidth;

        rData.mnActiveSheet = nActiveTab;
        rData.mnFirstVisSheet = nFirstVisTab;
        rData.mnTabBarWidth = nTabBarWidth;
        rData.mnVisibility = getFlagValue( nFlags, BIFF_WIN1_HIDDEN, XML_hidden, XML_visible );
        rData.mbMinimized = getFlag( nFlags, BIFF_WIN1_MINIMIZED );
        rData.mbShowHorScroll = getFlag( nFlags, BIFF_WIN1_SHOWHORSCROLL );
        rData.mbShowVerScroll = getFlag( nFlags, BIFF_WIN1_SHOWVERSCROLL );
        rData.mbShowTabBar = getFlag( nFlags, BIFF_WIN1_SHOWTABBAR );
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

