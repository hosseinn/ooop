/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: defnamesbuffer.cxx,v $
 *
 *  $Revision: 1.1.2.16 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/21 15:14:40 $
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

#include "oox/xls/defnamesbuffer.hxx"
#include <rtl/ustrbuf.hxx>
#include <com/sun/star/sheet/XNamedRanges.hpp>
#include <com/sun/star/sheet/XNamedRange.hpp>
#include <com/sun/star/sheet/XFormulaTokens.hpp>
#include <com/sun/star/sheet/XPrintAreas.hpp>
#include <com/sun/star/sheet/NamedRangeFlag.hpp>
#include "tokens.hxx"
#include "oox/core/attributelist.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/formulaparser.hxx"
#include "oox/xls/addressconverter.hxx"

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::container::XNameAccess;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::sheet::XNamedRanges;
using ::com::sun::star::sheet::XNamedRange;
using ::com::sun::star::sheet::XFormulaTokens;
using ::com::sun::star::sheet::XPrintAreas;
using ::oox::core::AttributeList;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

OoxDefinedNameData::OoxDefinedNameData() :
    mnSheet( -1 ),
    mbMacro( false ),
    mbFunction( false ),
    mbVBName( false ),
    mbHidden( false )
{
}

// ============================================================================

namespace {

const sal_Char* const spcLegacyPrefix = "Excel_BuiltIn_";
const sal_Char* const spcOoxPrefix = "_xlnm.";

const sal_Char* const sppcBaseNames[] =
{
    "Consolidate_Area", /* OOX */
    "Auto_Open",
    "Auto_Close",
    "Extract",          /* OOX */
    "Database",         /* OOX */
    "Criteria",         /* OOX */
    "Print_Area",       /* OOX */
    "Print_Titles",     /* OOX */
    "Recorder",
    "Data_Form",
    "Auto_Activate",
    "Auto_Deactivate",
    "Sheet_Title",      /* OOX */
    "_FilterDatabase"   /* OOX */
};

/** Localized names for _xlnm._FilterDatabase as used in BIFF5. */
const sal_Char* const sppcFilterDbNames[] =
{
    "_FilterDatabase",      // English
    "_FilterDatenbank"      // German
};

OUString lclGetBaseName( sal_Unicode cBuiltinId )
{
    OSL_ENSURE( cBuiltinId < STATIC_TABLE_SIZE( sppcBaseNames ), "lclGetBaseName - unknown builtin name" );
    OUStringBuffer aBuffer;
    if( cBuiltinId < STATIC_TABLE_SIZE( sppcBaseNames ) )
        aBuffer.appendAscii( sppcBaseNames[ cBuiltinId ] );
    else
        aBuffer.append( static_cast< sal_Int32 >( cBuiltinId ) );
    return aBuffer.makeStringAndClear();
}

OUString lclGetFinalName( sal_Unicode cBuiltinId )
{
    return OUStringBuffer().appendAscii( spcOoxPrefix ).append( lclGetBaseName( cBuiltinId ) ).makeStringAndClear();
}

sal_Unicode lclGetBuiltinIdFromOox( const OUString& rOoxName )
{
    OUString aPrefix = OUString::createFromAscii( spcOoxPrefix );
    sal_Int32 nPrefixLen = aPrefix.getLength();
    if( rOoxName.matchIgnoreAsciiCase( aPrefix ) )
    {
        for( sal_Unicode cBuiltinId = 0; cBuiltinId < STATIC_TABLE_SIZE( sppcBaseNames ); ++cBuiltinId )
        {
            OUString aBaseName = lclGetBaseName( cBuiltinId );
            sal_Int32 nBaseNameLen = aBaseName.getLength();
            if( (rOoxName.getLength() == nPrefixLen + nBaseNameLen) && rOoxName.matchIgnoreAsciiCase( aBaseName, nPrefixLen ) )
                return cBuiltinId;
        }
    }
    return BIFF_BUILTIN_UNKNOWN;
}

bool lclIsFilterDatabaseName( const OUString& rName )
{
    for( const sal_Char* const* ppcName = sppcFilterDbNames; ppcName < STATIC_TABLE_END( sppcFilterDbNames ); ++ppcName )
        if( rName.equalsIgnoreAsciiCaseAscii( *ppcName ) )
            return true;
    return false;
}

} // namespace

// ----------------------------------------------------------------------------

DefinedName::DefinedName( const GlobalDataHelper& rGlobalData, sal_Int32 nLocalSheet ) :
    GlobalDataHelper( rGlobalData ),
    mnTokenIndex( -1 ),
    mcBuiltinId( BIFF_BUILTIN_UNKNOWN ),
    mpStrm( 0 ),
    mnRecHandle( -1 ),
    mnRecPos( 0 ),
    mnFmlaSize( 0 )
{
    maOoxData.mnSheet = nLocalSheet;
}

void DefinedName::importDefinedName( const AttributeList& rAttribs )
{
    maOoxData.maName = rAttribs.getString( XML_name );
    maOoxData.mnSheet = rAttribs.getInteger( XML_localSheetId, -1 );
    maOoxData.mbMacro = rAttribs.getBool( XML_xlm, false );
    maOoxData.mbFunction = maOoxData.mbMacro && rAttribs.getBool( XML_function, false );
    maOoxData.mbVBName = maOoxData.mbMacro && rAttribs.getBool( XML_vbProcedure, false );
    maOoxData.mbHidden = rAttribs.getBool( XML_hidden, false );
    mcBuiltinId = lclGetBuiltinIdFromOox( maOoxData.maName );
}

void DefinedName::setFormula( const OUString& rFormula )
{
    maOoxData.maFormula = rFormula;
}

void DefinedName::importName( BiffInputStream& rStrm )
{
    BiffType eBiff = getBiff();
    sal_uInt16 nFlags = 0, nRefId = BIFF_NAME_GLOBAL, nTab = BIFF_NAME_GLOBAL;
    sal_uInt8 nNameLen = 0, nShortCut = 0;

    switch( eBiff )
    {
        case BIFF2:
        {
            sal_uInt8 nFlagsBiff2;
            rStrm >> nFlagsBiff2;
            rStrm.ignore( 1 );
            rStrm >> nShortCut >> nNameLen;
            mnFmlaSize = rStrm.readuInt8();
            setFlag( nFlags, BIFF_NAME_FUNC, getFlag( nFlagsBiff2, BIFF_NAME2_FUNC ) );
            maOoxData.maName = rStrm.readCharArray( nNameLen, getTextEncoding() );
        }
        break;
        case BIFF3:
        case BIFF4:
            rStrm >> nFlags >> nShortCut >> nNameLen >> mnFmlaSize;
            maOoxData.maName = rStrm.readCharArray( nNameLen, getTextEncoding() );
        break;
        case BIFF5:
            rStrm >> nFlags >> nShortCut >> nNameLen >> mnFmlaSize >> nRefId >> nTab;
            rStrm.ignore( 4 );
            maOoxData.maName = rStrm.readCharArray( nNameLen, getTextEncoding() );
        break;
        case BIFF8:
            rStrm >> nFlags >> nShortCut >> nNameLen >> mnFmlaSize >> nRefId >> nTab;
            rStrm.ignore( 4 );
            maOoxData.maName = rStrm.readUniString( nNameLen );
        break;
        case BIFF_UNKNOWN: break;
    }

    // macro function/command, hidden flag
    maOoxData.mbMacro = getFlag( nFlags, BIFF_NAME_MACRO );
    maOoxData.mbFunction = maOoxData.mbMacro && getFlag( nFlags, BIFF_NAME_FUNC );
    maOoxData.mbVBName = maOoxData.mbMacro && getFlag( nFlags, BIFF_NAME_VBNAME );
    maOoxData.mbHidden = getFlag( nFlags, BIFF_NAME_HIDDEN );

    // get builtin name index from name
    if( getFlag( nFlags, BIFF_NAME_BUILTIN ) )
    {
        OSL_ENSURE( maOoxData.maName.getLength() == 1, "DefinedName::importName - wrong builtin name" );
        if( maOoxData.maName.getLength() > 0 )
        {
            mcBuiltinId = maOoxData.maName[ 0 ];
            if( mcBuiltinId == '?' )      // the NUL character is imported as '?'
                mcBuiltinId = BIFF_BUILTIN_CONSOLIDATEAREA;
        }
    }
    /*  In BIFF5, _xlnm._FilterDatabase appears as hidden user name without
        built-in flag, and even worse, localized. */
    else if( (eBiff == BIFF5) && lclIsFilterDatabaseName( maOoxData.maName ) )
    {
        mcBuiltinId = BIFF_BUILTIN_FILTERDATABASE;
    }

    // unhide built-in names (_xlnm._FilterDatabase is always hidden)
    if( isBuiltinName() )
        maOoxData.mbHidden = false;

    // get sheet index for local names in BIFF5-BIFF8
    switch( getBiff() )
    {
        case BIFF2:
        case BIFF3:
        case BIFF4:
        break;
        case BIFF5:
            // #i44019# nTab may be invalid, TODO: resolve nRefId to sheet index
            if( nRefId != BIFF_NAME_GLOBAL )
                maOoxData.mnSheet = nRefId - 1;
        break;
        case BIFF8:
            if( nTab != BIFF_NAME_GLOBAL )
                maOoxData.mnSheet = nTab - 1;
        break;
        case BIFF_UNKNOWN:
        break;
    }

    // store record position to be able to import token array later
    mpStrm = &rStrm;
    mnRecHandle = rStrm.getRecHandle();
    mnRecPos = rStrm.getRecPos();
}

void DefinedName::createNameObject()
{
    // do not create hidden names and names for (macro) functions
    if( maOoxData.mbHidden || maOoxData.mbFunction )
        return;

    // convert original name to final Calc name
    if( maOoxData.mbVBName )
        maFinalName = maOoxData.maName;
    else if( isBuiltinName() )
        maFinalName = lclGetFinalName( mcBuiltinId );
    else
        maFinalName = maOoxData.maName;         //! TODO convert to valid name

    // append sheet index for local names in multi-sheet documents
    if( isWorkbookFile() && !isGlobalName() )
        maFinalName = OUStringBuffer( maFinalName ).append( sal_Unicode( '.' ) ).append( maOoxData.mnSheet + 1 ).makeStringAndClear();

    // special flags for this name
    sal_Int32 nNameFlags = 0;
    using namespace ::com::sun::star::sheet::NamedRangeFlag;
    if( !isGlobalName() ) switch( mcBuiltinId )
    {
        case BIFF_BUILTIN_CRITERIA:     nNameFlags = FILTER_CRITERIA;               break;
        case BIFF_BUILTIN_PRINTAREA:    nNameFlags = PRINT_AREA;                    break;
        case BIFF_BUILTIN_PRINTTITLES:  nNameFlags = COLUMN_HEADER | ROW_HEADER;    break;
    }

    // create the name and insert it into the document, maFinalName will be changed to the resulting name
    mxNamedRange = getDefinedNames().createDefinedName( maFinalName, nNameFlags );
    // index of this defined name used in formula token arrays
    mnTokenIndex = getDefinedNames().getTokenIndex( mxNamedRange );
}

void DefinedName::convertFormula()
{
    Reference< XFormulaTokens > xTokens( mxNamedRange, UNO_QUERY );
    if( xTokens.is() )
    {
        // convert and set formula of the defined name
        SimpleFormulaContext aContext( xTokens );
        switch( getFilterType() )
        {
            case FILTER_OOX:
                if( maOoxData.maFormula.getLength() == 0 )
                    getFormulaParser().convertErrorToFormula( aContext, BIFF_ERR_NAME );
                else
                    getFormulaParser().importFormula( aContext, maOoxData.maFormula );
            break;
            case FILTER_BIFF:
            {
                if( mnFmlaSize == 0 )
                {
                    getFormulaParser().convertErrorToFormula( aContext, BIFF_ERR_NAME );
                }
                else
                {
                    OSL_ENSURE( mpStrm, "DefinedName::convertFormula - missing BIFF stream" );
                    sal_Int64 nCurrRecHandle = mpStrm->getRecHandle();
                    if( mpStrm->startRecordByHandle( mnRecHandle ) )
                    {
                        mpStrm->seek( mnRecPos );
                        getFormulaParser().importFormula( aContext, *mpStrm, &mnFmlaSize );
                    }
                    mpStrm->startRecordByHandle( nCurrRecHandle );
                }
            }
            break;
            case FILTER_UNKNOWN: break;
        }

        // set builtin names (print ranges, repeated titles, filter ranges)
        if( !isGlobalName() ) switch( mcBuiltinId )
        {
            case BIFF_BUILTIN_PRINTAREA:
            {
                Reference< XPrintAreas > xPrintAreas( getSheet( maOoxData.mnSheet ), UNO_QUERY );
                Sequence< CellRangeAddress > aPrintRanges =
                    getFormulaParser().convertToCellRangeList( xTokens->getTokens(), maOoxData.mnSheet );
                if( xPrintAreas.is() && aPrintRanges.hasElements() )
                    xPrintAreas->setPrintAreas( aPrintRanges );
            }
            break;
            case BIFF_BUILTIN_PRINTTITLES:
            {
                Reference< XPrintAreas > xPrintAreas( getSheet( maOoxData.mnSheet ), UNO_QUERY );
                Sequence< CellRangeAddress > aTitleRanges =
                    getFormulaParser().convertToCellRangeList( xTokens->getTokens(), maOoxData.mnSheet );
                if( xPrintAreas.is() && aTitleRanges.hasElements() )
                {
                    bool bHasRowTitles = false;
                    bool bHasColTitles = false;
                    const CellAddress& rMaxPos = getAddressConverter().getMaxAddress();
                    const CellRangeAddress* pTitleRange = aTitleRanges.getConstArray();
                    const CellRangeAddress* pTitleRangeEnd = pTitleRange + aTitleRanges.getLength();
                    for( ; (pTitleRange != pTitleRangeEnd) && (!bHasRowTitles || !bHasColTitles); ++pTitleRange )
                    {
                        bool bFullRow = (pTitleRange->StartColumn == 0) && (pTitleRange->EndColumn >= rMaxPos.Column);
                        bool bFullCol = (pTitleRange->StartRow == 0) && (pTitleRange->EndRow >= rMaxPos.Row);
                        if( !bHasRowTitles && bFullRow && !bFullCol )
                        {
                            xPrintAreas->setTitleRows( *pTitleRange );
                            xPrintAreas->setPrintTitleRows( sal_True );
                            bHasRowTitles = true;
                        }
                        else if( !bHasColTitles && bFullCol && !bFullRow )
                        {
                            xPrintAreas->setTitleColumns( *pTitleRange );
                            xPrintAreas->setPrintTitleColumns( sal_True );
                            bHasColTitles = true;
                        }
                    }
                }
            }
            break;
        }
    }
}

// ============================================================================

DefinedNamesBuffer::DefinedNamesBuffer( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData ),
    maTokenIndexProp( CREATE_OUSTRING( "TokenIndex" ) ),
    mnLocalSheet( -1 )
{
}

Reference< XNamedRange > DefinedNamesBuffer::createDefinedName( OUString& rName, sal_Int32 nNameFlags ) const
{
    // find an unused name
    Reference< XNamedRanges > xNamedRanges = getNamedRanges();
    Reference< XNameAccess > xNameAccess( xNamedRanges, UNO_QUERY );
    if( xNameAccess.is() )
        rName = ContainerHelper::getUnusedName( xNameAccess, rName, '.' );

    // create the name and insert it into the Calc document
    Reference< XNamedRange > xNamedRange;
    if( xNamedRanges.is() && (rName.getLength() > 0) ) try
    {
        xNamedRanges->addNewByName( rName, OUString(), CellAddress( 0, 0, 0 ), nNameFlags );
        xNamedRange.set( xNamedRanges->getByName( rName ), UNO_QUERY );
    }
    catch( Exception& )
    {
        OSL_ENSURE( false, "DefinedNamesBuffer::createDefinedName - cannot create defined name" );
    }
    return xNamedRange;
}

sal_Int32 DefinedNamesBuffer::getTokenIndex( const Reference< XNamedRange >& rxNamedRange ) const
{
    PropertySet aPropSet( rxNamedRange );
    sal_Int32 nIndex = -1;
    return aPropSet.getProperty( nIndex, maTokenIndexProp ) ? nIndex : -1;
}

void DefinedNamesBuffer::setLocalSheetIndex( sal_Int32 nLocalSheet )
{
    mnLocalSheet = nLocalSheet;
}

DefinedNameRef DefinedNamesBuffer::importDefinedName( const AttributeList& rAttribs )
{
    DefinedNameRef xDefName = createDefinedName();
    xDefName->importDefinedName( rAttribs );
    return xDefName;
}

void DefinedNamesBuffer::importName( BiffInputStream& rStrm )
{
    DefinedNameRef xDefName = createDefinedName();
    xDefName->importName( rStrm );
}

void DefinedNamesBuffer::finalizeImport()
{
    /*  First insert all names without formula definition into the document. */
    maDefNames.forEachMem( &DefinedName::createNameObject );
    /*  Now convert all name formulas, so that the formula parser can find all
        names in case of circular dependencies. */
    maDefNames.forEachMem( &DefinedName::convertFormula );
}

DefinedNameRef DefinedNamesBuffer::getByBiffIndex( sal_uInt16 nIndex ) const
{
    OSL_ENSURE( nIndex > 0, "DefinedNamesBuffer::getByBiffIndex - invalid name index" );
    return maDefNames.get( static_cast< sal_Int32 >( nIndex ) - 1 );
}

DefinedNameRef DefinedNamesBuffer::getByOoxName( const OUString& rOoxName, sal_Int32 nSheet ) const
{
    DefinedNameRef xGlobalName;   // a found global name
    DefinedNameRef xLocalName;    // a found local name
    for( DefNameVec::const_iterator aIt = maDefNames.begin(), aEnd = maDefNames.end(); (aIt != aEnd) && !xLocalName; ++aIt )
    {
        DefinedNameRef xCurrName = *aIt;
        if( xCurrName->getOoxName() == rOoxName )
        {
            if( xCurrName->getSheetIndex() == nSheet )
                xLocalName = xCurrName;
            else if( xCurrName->isGlobalName() )
                xGlobalName = xCurrName;
        }
    }
    return xLocalName.get() ? xLocalName : xGlobalName;
}

DefinedNameRef DefinedNamesBuffer::createDefinedName()
{
    DefinedNameRef xDefName( new DefinedName( getGlobalData(), mnLocalSheet ) );
    maDefNames.push_back( xDefName );
    return xDefName;
}

// ============================================================================

} // namespace xls
} // namespace oox

