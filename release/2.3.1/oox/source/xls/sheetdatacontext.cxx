/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sheetdatacontext.cxx,v $
 *
 *  $Revision: 1.1.2.60 $
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

#include "oox/xls/sheetdatacontext.hxx"
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/table/XCellRange.hpp>
#include <com/sun/star/sheet/XFormulaTokens.hpp>
#include <com/sun/star/sheet/XArrayFormulaTokens.hpp>
#include <com/sun/star/text/XText.hpp>
#include "oox/core/propertyset.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/formulaparser.hxx"
#include "oox/xls/pivottablebuffer.hxx"
#include "oox/xls/richstringcontext.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/unitconverter.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::table::XCell;
using ::com::sun::star::table::XCellRange;
using ::com::sun::star::sheet::XFormulaTokens;
using ::com::sun::star::sheet::XArrayFormulaTokens;
using ::com::sun::star::text::XText;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::oox::core::AttributeList;

namespace oox {
namespace xls {

// ============================================================================

namespace {

/** Formula context that supports shared formulas and table operations. */
class CellFormulaContext : public SimpleFormulaContext, public WorksheetHelper
{
public:
    explicit            CellFormulaContext(
                            const Reference< XFormulaTokens >& rxTokens,
                            const CellAddress& rCellPos,
                            const WorksheetHelper& rSheetHelper );

    virtual void        setSharedFormula( sal_Int32 nSharedFmlaId );
    virtual void        setTableOperation( sal_Int32 nTableOpId );
};

CellFormulaContext::CellFormulaContext( const Reference< XFormulaTokens >& rxTokens,
        const CellAddress& rCellPos, const WorksheetHelper& rSheetHelper ) :
    SimpleFormulaContext( rxTokens ),
    WorksheetHelper( rSheetHelper )
{
    setBaseAddress( rCellPos, false );
}

void CellFormulaContext::setSharedFormula( sal_Int32 nSharedId )
{
    if( nSharedId >= 0 )
    {
        sal_Int32 nTokenIndex = getSharedFormulas().getTokenIndexFromId( nSharedId );
        getFormulaParser().convertNameToFormula( *this, nTokenIndex );
    }
}

void CellFormulaContext::setTableOperation( sal_Int32 nTableOpId )
{
    if( nTableOpId >= 0 )
    {
        getFormulaParser().convertErrorToFormula( *this, BIFF_ERR_NA );
    }
}

} // namespace

// ============================================================================

OoxSheetDataContext::OoxSheetDataContext( const WorksheetFragmentBase& rFragment ) :
    WorksheetContextBase( rFragment )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxSheetDataContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sheetData ):
            return  (nElement == XLS_TOKEN( row ));
        case XLS_TOKEN( row ):
            return  (nElement == XLS_TOKEN( c ));
        case XLS_TOKEN( c ):
            return  maCurrCell.mxCell.is() &&
                   ((nElement == XLS_TOKEN( v )) ||
                    (nElement == XLS_TOKEN( is )) ||
                    (nElement == XLS_TOKEN( f )));
    }
    return false;
}

Reference< XFastContextHandler > OoxSheetDataContext::onCreateContext( sal_Int32 nElement, const AttributeList& /*rAttribs*/ )
{
    switch( nElement )
    {
        case XLS_TOKEN( is ):
            mxInlineStr.reset( new RichString( getGlobalData() ) );
            return new OoxRichStringContext( *this, mxInlineStr );
    }
    return this;
}

void OoxSheetDataContext::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( row ):
            importRow( rAttribs );
        break;
        case XLS_TOKEN( c ):
            importCell( rAttribs );
        break;
        case XLS_TOKEN( f ):
            importFormula( rAttribs );
        break;
    }
}

void OoxSheetDataContext::onEndElement( const OUString& rChars )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( v ):
            maCurrCell.maValueStr = rChars;
            maCurrCell.mbHasValueStr = true;
        break;
        case XLS_TOKEN( f ):
            maCurrCell.maFormulaStr = rChars;
        break;

        case XLS_TOKEN( c ):
        {
            if( maCurrCell.mxCell.is() )
            {
                try
                {
                    if( maCurrCell.mnFormulaType != XML_TOKEN_INVALID )
                    {
                        switch( maCurrCell.mnFormulaType )
                        {
                            case XML_normal:
                            {
                                Reference< XFormulaTokens > xTokens( maCurrCell.mxCell, UNO_QUERY_THROW );
                                SimpleFormulaContext aContext( xTokens );
                                aContext.setBaseAddress( maCurrCell.maAddress );
                                getFormulaParser().importFormula( aContext, maCurrCell.maFormulaStr );
                            }
                            break;

                            case XML_array:
                                if( (maCurrCell.maFormulaRef.getLength() > 0) && (maCurrCell.maFormulaStr.getLength() > 0) )
                                {
                                    CellRangeAddress aArrayRange;
                                    Reference< XArrayFormulaTokens > xTokens( getCellRange( maCurrCell.maFormulaRef, &aArrayRange ), UNO_QUERY_THROW );
                                    ArrayFormulaContext aContext( xTokens, aArrayRange );
                                    getFormulaParser().importFormula( aContext, maCurrCell.maFormulaStr );
                                }
                            break;

                            case XML_shared:
                            {
                                if( (maCurrCell.mnSharedId >= 0) && (maCurrCell.maFormulaStr.getLength() > 0) )
                                    getSharedFormulas().importSharedFmla( maCurrCell.maFormulaStr, maCurrCell.mnSharedId, maCurrCell.maAddress );
                                Reference< XFormulaTokens > xTokens( maCurrCell.mxCell, UNO_QUERY_THROW );
                                CellFormulaContext aContext( xTokens, maCurrCell.maAddress, getWorksheetHelper() );
                                aContext.setSharedFormula( maCurrCell.mnSharedId );
                            }
                            break;
                        }
                    }
                    else if( maCurrCell.mbHasValueStr )
                    {
                        // implemented in WorksheetHelper class
                        setCell( maCurrCell );
                    }
                    else if( (maCurrCell.mnCellType == XML_inlineStr) && mxInlineStr.get() )
                    {
                        // convert font settings
                        mxInlineStr->finalizeImport();
                        // write string to cell
                        Reference< XText > xText( maCurrCell.mxCell, UNO_QUERY_THROW );
                        mxInlineStr->convert( xText, maCurrCell.mnXfId );
                    }
                    else
                    {
                        // empty cell, update cell type
                        maCurrCell.mnCellType = XML_TOKEN_INVALID;
                    }
                }
                catch( Exception& )
                {
                }

                // store the cell formatting data
                setCellFormat( maCurrCell );
            }
        }
        break;
    }
}

// private --------------------------------------------------------------------

void OoxSheetDataContext::importRow( const AttributeList& rAttribs )
{
    OoxRowData aData;
    aData.mnFirstRow = aData.mnLastRow = rAttribs.getInteger( XML_r, -1 );
    aData.mfHeight = rAttribs.getDouble( XML_ht, 0.0 );
    aData.mnXfId = rAttribs.getInteger( XML_s, -1 );
    aData.mnLevel = rAttribs.getInteger( XML_outlineLevel, 0 );
    aData.mbFormatted = rAttribs.getBool( XML_customFormat, false );
    aData.mbHidden = rAttribs.getBool( XML_hidden, false );
    aData.mbCollapsed = rAttribs.getBool( XML_collapsed, false );
    // set row properties in the current sheet
    setRowData( aData );
}

void OoxSheetDataContext::importCell( const AttributeList& rAttribs )
{
    maCurrCell.reset();
    maCurrCell.mxCell = getCell( rAttribs.getString( XML_r ), &maCurrCell.maAddress );
    maCurrCell.mnCellType = rAttribs.getToken( XML_t, XML_n );
    maCurrCell.mnXfId = rAttribs.getInteger( XML_s, -1 );
    mxInlineStr.reset();

    if( maCurrCell.mxCell.is() && getPivotTables().isOverlapping( maCurrCell.maAddress ) )
        // This cell overlaps a pivot table.  Skip it.
        maCurrCell.mxCell.clear();
}

void OoxSheetDataContext::importFormula( const AttributeList& rAttribs )
{
    maCurrCell.maFormulaRef = rAttribs.getString( XML_ref );
    maCurrCell.mnFormulaType = rAttribs.getToken( XML_t, XML_normal );
    maCurrCell.mnSharedId = rAttribs.getInteger( XML_si, -1 );
}

// ============================================================================

OoxExternalSheetDataContext::OoxExternalSheetDataContext(
        const GlobalFragmentBase& rFragment, const WorksheetHelper& rSheetHelper ) :
    WorksheetContextBase( rFragment, rSheetHelper )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxExternalSheetDataContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( sheetData ):
            return  (nElement == XLS_TOKEN( row ));
        case XLS_TOKEN( row ):
            return  (nElement == XLS_TOKEN( cell ));
        case XLS_TOKEN( cell ):
            return  (nElement == XLS_TOKEN( v )) && maCurrCell.mxCell.is();
    }
    return false;
}

void OoxExternalSheetDataContext::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( cell ):
            importCell( rAttribs );
        break;
    }
}

void OoxExternalSheetDataContext::onEndElement( const OUString& rChars )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( v ):
            maCurrCell.maValueStr = rChars;
            maCurrCell.mbHasValueStr = true;
        break;

        case XLS_TOKEN( cell ):
            if( maCurrCell.mxCell.is() && maCurrCell.mbHasValueStr )
                setCell( maCurrCell );
        break;
    }
}

// private --------------------------------------------------------------------

void OoxExternalSheetDataContext::importCell( const AttributeList& rAttribs )
{
    maCurrCell.reset();
    maCurrCell.mxCell = getCell( rAttribs.getString( XML_r ), &maCurrCell.maAddress );
    maCurrCell.mnCellType = rAttribs.getToken( XML_t, XML_n );
}

// ============================================================================

BiffSheetDataContext::BiffSheetDataContext( const WorksheetHelper& rSheetHelper ) :
    WorksheetHelper( rSheetHelper ),
    mnBiff2XfId( 0 )
{
    mnArrayIgnoreSize = (getBiff() == BIFF2) ? 1 : ((getBiff() <= BIFF4) ? 2 : 6);
    switch( getBiff() )
    {
        case BIFF2:
            mnFormulaIgnoreSize = 9;    // double formula result, 1 byte flags
            mnArrayIgnoreSize = 1;      // recalc-always flag
        break;
        case BIFF3:
        case BIFF4:
            mnFormulaIgnoreSize = 10;   // double formula result, 2 byte flags
            mnArrayIgnoreSize = 2;      // 2 byte flags
        break;
        case BIFF5:
        case BIFF8:
            mnFormulaIgnoreSize = 14;   // double formula result, 2 byte flags, 4 bytes nothing
            mnArrayIgnoreSize = 6;      // 2 byte flags, 4 bytes nothing
        break;
        case BIFF_UNKNOWN: break;
    }
}

void BiffSheetDataContext::importRecord( BiffInputStream& rStrm )
{
    sal_uInt16 nRecId = rStrm.getRecId();
    switch( nRecId )
    {
        // records in all BIFF versions
        case BIFF2_ID_ARRAY:        // #i72713#
        case BIFF3_ID_ARRAY:        importArray( rStrm );       break;
        case BIFF2_ID_BLANK:
        case BIFF3_ID_BLANK:        importBlank( rStrm );       break;
        case BIFF2_ID_BOOLERR:
        case BIFF3_ID_BOOLERR:      importBoolErr( rStrm );     break;
        case BIFF2_ID_INTEGER:      importInteger( rStrm );     break;
        case BIFF_ID_IXFE:          rStrm >> mnBiff2XfId;       break;
        case BIFF2_ID_LABEL:
        case BIFF3_ID_LABEL:        importLabel( rStrm );       break;
        case BIFF2_ID_NUMBER:
        case BIFF3_ID_NUMBER:       importNumber( rStrm );      break;
        case BIFF_ID_RK:            importRk( rStrm );          break;

        // BIFF specific records
        default: switch( getBiff() )
        {
            case BIFF2: switch( nRecId )
            {
                case BIFF2_ID_FORMULA:      importFormula( rStrm );     break;
                case BIFF2_ID_ROW:          importRow( rStrm );         break;
            }
            break;

            case BIFF3: switch( nRecId )
            {
                case BIFF3_ID_FORMULA:      importFormula( rStrm );     break;
                case BIFF3_ID_ROW:          importRow( rStrm );         break;
            }
            break;

            case BIFF4: switch( nRecId )
            {
                case BIFF4_ID_FORMULA:      importFormula( rStrm );     break;
                case BIFF3_ID_ROW:          importRow( rStrm );         break;
            }
            break;

            case BIFF5: switch( nRecId )
            {
                case BIFF3_ID_FORMULA:
                case BIFF4_ID_FORMULA:
                case BIFF5_ID_FORMULA:      importFormula( rStrm );     break;
                case BIFF_ID_MULTBLANK:     importMultBlank( rStrm );   break;
                case BIFF_ID_MULTRK:        importMultRk( rStrm );      break;
                case BIFF3_ID_ROW:          importRow( rStrm );         break;
                case BIFF_ID_RSTRING:       importLabel( rStrm );       break;
            }
            break;

            case BIFF8: switch( nRecId )
            {
                case BIFF3_ID_FORMULA:
                case BIFF4_ID_FORMULA:
                case BIFF5_ID_FORMULA:      importFormula( rStrm );     break;
                case BIFF_ID_LABELSST:      importLabelSst( rStrm );    break;
                case BIFF_ID_MULTBLANK:     importMultBlank( rStrm );   break;
                case BIFF_ID_MULTRK:        importMultRk( rStrm );      break;
                case BIFF3_ID_ROW:          importRow( rStrm );         break;
                case BIFF_ID_RSTRING:       importLabel( rStrm );       break;
            }
            break;

            case BIFF_UNKNOWN: break;
        }
    }
}

// private --------------------------------------------------------------------

void BiffSheetDataContext::setCurrCell( const BiffAddress& rBiffAddr )
{
    maCurrCell.reset();
    maCurrCell.mxCell = getCell( rBiffAddr, &maCurrCell.maAddress );
}

void BiffSheetDataContext::importXfId( BiffInputStream& rStrm, bool bBiff2 )
{
    if( bBiff2 )
    {
        sal_uInt8 nBiff2XfId;
        rStrm >> nBiff2XfId;
        rStrm.ignore( 2 );
        maCurrCell.mnXfId = nBiff2XfId & BIFF2_XF_MASK;
        if( maCurrCell.mnXfId == BIFF_IXFE_USE_CACHED )
            maCurrCell.mnXfId = mnBiff2XfId;
    }
    else
    {
        maCurrCell.mnXfId = rStrm.readuInt16();
    }
}

void BiffSheetDataContext::importCellHeader( BiffInputStream& rStrm, bool bBiff2 )
{
    BiffAddress aBiffAddr;
    rStrm >> aBiffAddr;
    setCurrCell( aBiffAddr );
    importXfId( rStrm, bBiff2 );
}

void BiffSheetDataContext::importBlank( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, rStrm.getRecId() == BIFF2_ID_BLANK );
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importBoolErr( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, rStrm.getRecId() == BIFF2_ID_BOOLERR );
    if( maCurrCell.mxCell.is() )
    {
        sal_uInt8 nValue, nType;
        rStrm >> nValue >> nType;
        switch( nType )
        {
            case BIFF_BOOLERR_BOOL:
                maCurrCell.mnCellType = XML_e;
                setBooleanCell( maCurrCell.mxCell, nValue != 0 );
                // #108770# set 'Standard' number format for all Boolean cells
                maCurrCell.mnNumFmtId = 0;
            break;
            case BIFF_BOOLERR_ERROR:
                maCurrCell.mnCellType = XML_b;
                setErrorCell( maCurrCell.mxCell, nValue );
            break;
            default:
                OSL_ENSURE( false, "BiffSheetDataContext::importBoolErr - unknown cell type" );
        }
    }
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importFormula( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, getBiff() == BIFF2 );
    maCurrCell.mnCellType = XML_n;
    Reference< XFormulaTokens > xTokens( maCurrCell.mxCell, UNO_QUERY );
    if( xTokens.is() )
    {
        rStrm.ignore( mnFormulaIgnoreSize );

        // shared formula (Excel ignores the BIFF_FORMULA_SHARED flag)
        if( getBiff() >= BIFF5 )
        {
            // try to read shared formula definition from SHRFMLA, before creating the first cell
            if( rStrm.getNextRecId() == BIFF_ID_SHRFMLA )
            {
                sal_Int64 nRecHandle = rStrm.getRecHandle();
                sal_uInt32 nRecPos = rStrm.getRecPos();
                if( rStrm.startNextRecord() )
                    getSharedFormulas().importShrFmla( rStrm, maCurrCell.maAddress );
                rStrm.startRecordByHandle( nRecHandle );
                rStrm.seek( nRecPos );
            }
            // read the formula
            CellFormulaContext aContext( xTokens, maCurrCell.maAddress, getWorksheetHelper() );
            getFormulaParser().importFormula( aContext, rStrm );
        }
        else
        {
            SimpleFormulaContext aContext( xTokens );
            aContext.setBaseAddress( maCurrCell.maAddress );
            getFormulaParser().importFormula( aContext, rStrm );
        }
    }
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importInteger( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, true );
    maCurrCell.mnCellType = XML_n;
    if( maCurrCell.mxCell.is() )
        maCurrCell.mxCell->setValue( rStrm.readuInt16() );
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importLabel( BiffInputStream& rStrm )
{
    bool bBiff2Xf = rStrm.getRecId() == BIFF2_ID_LABEL;
    importCellHeader( rStrm, bBiff2Xf );
    maCurrCell.mnCellType = XML_inlineStr;
    Reference< XText > xText( maCurrCell.mxCell, UNO_QUERY );
    if( xText.is() )
    {
        /*  the deep secrets of BIFF type and record identifier...
            record id   BIFF    XF type     String type
            0x0004      2-7     3 byte      8-bit length, byte string
            0x0004      8       3 byte      16-bit length, unicode string
            0x0204      2-7     2 byte      16-bit length, byte string
            0x0204      8       2 byte      16-bit length, unicode string */

        RichString aString( getGlobalData() );
        if( getBiff() == BIFF8 )
        {
            aString.importUniString( rStrm );
        }
        else
        {
            // #i63105# use text encoding from FONT record
            rtl_TextEncoding eTextEnc = getTextEncoding();
            if( const Font* pFont = getStyles().getFontFromCellXf( maCurrCell.mnXfId ).get() )
                eTextEnc = pFont->getFontEncoding();
            BiffStringFlags nFlags = bBiff2Xf ? BIFF_STR_8BITLENGTH : BIFF_STR_DEFAULT;
            setFlag( nFlags, BIFF_STR_EXTRAFONTS, rStrm.getRecId() == BIFF_ID_RSTRING );
            aString.importByteString( rStrm, eTextEnc, nFlags );
        }
        aString.finalizeImport();
        aString.convert( xText, maCurrCell.mnXfId );
    }
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importLabelSst( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, false );
    maCurrCell.mnCellType = XML_s;
    Reference< XText > xText( maCurrCell.mxCell, UNO_QUERY );
    if( xText.is() )
        getSharedStrings().convertString( xText, rStrm.readInt32(), maCurrCell.mnXfId );
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importMultBlank( BiffInputStream& rStrm )
{
    BiffAddress aBiffAddr;
    for( rStrm >> aBiffAddr; rStrm.getRecLeft() > 2; ++aBiffAddr.mnCol )
    {
        setCurrCell( aBiffAddr );
        importXfId( rStrm, false );
        setCellFormat( maCurrCell );
    }
}

void BiffSheetDataContext::importMultRk( BiffInputStream& rStrm )
{
    BiffAddress aBiffAddr;
    for( rStrm >> aBiffAddr; rStrm.getRecLeft() > 2; ++aBiffAddr.mnCol )
    {
        setCurrCell( aBiffAddr );
        maCurrCell.mnCellType = XML_n;
        importXfId( rStrm, false );
        sal_Int32 nRkValue = rStrm.readInt32();
        if( maCurrCell.mxCell.is() )
            maCurrCell.mxCell->setValue( BiffHelper::calcDoubleFromRk( nRkValue ) );
        setCellFormat( maCurrCell );
    }
}

void BiffSheetDataContext::importNumber( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, rStrm.getRecId() == BIFF2_ID_NUMBER );
    maCurrCell.mnCellType = XML_n;
    if( maCurrCell.mxCell.is() )
        maCurrCell.mxCell->setValue( rStrm.readDouble() );
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importRk( BiffInputStream& rStrm )
{
    importCellHeader( rStrm, false );
    maCurrCell.mnCellType = XML_n;
    if( maCurrCell.mxCell.is() )
        maCurrCell.mxCell->setValue( BiffHelper::calcDoubleFromRk( rStrm.readInt32() ) );
    setCellFormat( maCurrCell );
}

void BiffSheetDataContext::importRow( BiffInputStream& rStrm )
{
    sal_uInt16 nRow, nHeight, nFlags = 0, nXfId = 0;
    bool bFormatted = false;

    rStrm >> nRow;
    rStrm.ignore( 4 );
    rStrm >> nHeight;
    if( getBiff() == BIFF2 )
    {
        rStrm.ignore( 2 );
        bFormatted = rStrm.readuInt8() == BIFF2_ROW_FORMATTED;
        if( bFormatted )
        {
            rStrm.ignore( 3 );
            rStrm >> nXfId;
        }
    }
    else
    {
        rStrm.ignore( 4 );
        rStrm >> nFlags >> nXfId;
    }

    OoxRowData aData;
    // row index is 0-based in BIFF, but OoxRowData expects 1-based
    aData.mnFirstRow = aData.mnLastRow = nRow + 1;
    // row height is in twips in BIFF, convert to points
    aData.mfHeight = (nHeight & BIFF_ROW_HEIGHTMASK) / 20.0;
    aData.mnXfId = nXfId & BIFF_ROW_XFMASK;
    aData.mnLevel = extractValue< sal_Int32 >( nFlags, 0, 3 );
    aData.mbFormatted = bFormatted;
    aData.mbHidden = getFlag( nFlags, BIFF_ROW_HIDDEN );
    aData.mbCollapsed = getFlag( nFlags, BIFF_ROW_COLLAPSED );
    // set row properties in the current sheet
    setRowData( aData );
}

void BiffSheetDataContext::importArray( BiffInputStream& rStrm )
{
    BiffRange aBiffRange;
    aBiffRange.read( rStrm, false );
    CellRangeAddress aArrayRange;
    Reference< XCellRange > xRange = getCellRange( aBiffRange, &aArrayRange );
    Reference< XArrayFormulaTokens > xTokens( xRange, UNO_QUERY );
    if( xRange.is() && xTokens.is() )
    {
        rStrm.ignore( mnArrayIgnoreSize );
        ArrayFormulaContext aContext( xTokens, aArrayRange );
        getFormulaParser().importFormula( aContext, rStrm );
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

