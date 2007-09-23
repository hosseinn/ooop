/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: formulaparser.cxx,v $
 *
 *  $Revision: 1.1.2.53 $
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

#include "oox/xls/formulaparser.hxx"
#include <com/sun/star/sheet/FormulaToken.hpp>
#include <com/sun/star/sheet/ReferenceFlags.hpp>
#include <com/sun/star/sheet/SingleReference.hpp>
#include <com/sun/star/sheet/ComplexReference.hpp>
#include <com/sun/star/sheet/XFormulaParser.hpp>
#include <com/sun/star/sheet/XFormulaTokens.hpp>
#include <com/sun/star/sheet/XMultiFormulaTokens.hpp>
#include <com/sun/star/sheet/XArrayFormulaTokens.hpp>
#include "oox/core/containerhelper.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/externallinkbuffer.hxx"
#include "oox/xls/defnamesbuffer.hxx"
#include "oox/xls/worksheethelper.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::sheet::SingleReference;
using ::com::sun::star::sheet::ComplexReference;
using ::com::sun::star::sheet::XFormulaParser;
using ::com::sun::star::sheet::XFormulaTokens;
using ::com::sun::star::sheet::XMultiFormulaTokens;
using ::com::sun::star::sheet::XArrayFormulaTokens;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySet;
using namespace ::com::sun::star::sheet::ReferenceFlags;

namespace oox {
namespace xls {

// ============================================================================

FormulaContext::FormulaContext() :
    maBaseAddress( 0, 0, 0 ),
    mbRelativeAsOffset( true )
{
}

FormulaContext::~FormulaContext()
{
}

void FormulaContext::setBaseAddress( const CellAddress& rBaseAddress, bool bRelativeAsOffset )
{
    maBaseAddress = rBaseAddress;
    mbRelativeAsOffset = bRelativeAsOffset;
}

void FormulaContext::setSharedFormula( sal_Int32 )
{
}

void FormulaContext::setTableOperation( sal_Int32 )
{
}

// ----------------------------------------------------------------------------

AnyFormulaContext::AnyFormulaContext( Any& rAny ) :
    mrAny( rAny )
{
}

void AnyFormulaContext::setTokens( const ApiTokenSequence& rTokens )
{
    mrAny <<= rTokens;
}

// ----------------------------------------------------------------------------

SimpleFormulaContext::SimpleFormulaContext( const Reference< XFormulaTokens >& rxTokens ) :
    mxTokens( rxTokens )
{
    OSL_ENSURE( mxTokens.is(), "SimpleFormulaContext::SimpleFormulaContext - missing XFormulaTokens interface" );
}

void SimpleFormulaContext::setTokens( const ApiTokenSequence& rTokens )
{
    mxTokens->setTokens( rTokens );
}

// ----------------------------------------------------------------------------

MultiFormulaContext::MultiFormulaContext( const Reference< XMultiFormulaTokens >& rxTokens, sal_Int32 nIndex ) :
    mxTokens( rxTokens ),
    mnIndex( nIndex )
{
    OSL_ENSURE( mxTokens.is(), "MultiFormulaContext::MultiFormulaContext - missing XMultiFormulaTokens interface" );
}

void MultiFormulaContext::setTokens( const ApiTokenSequence& rTokens )
{
    try
    {
        mxTokens->setTokens( mnIndex, rTokens );
    }
    catch( const Exception& )
    {
        OSL_ENSURE( false, "MultiFormulaContext::setTokens - index out of range" );
    }
}

// ----------------------------------------------------------------------------

ArrayFormulaContext::ArrayFormulaContext(
        const Reference< XArrayFormulaTokens >& rxTokens, const CellRangeAddress& rArrayRange ) :
    mxTokens( rxTokens )
{
    OSL_ENSURE( mxTokens.is(), "ArrayFormulaContext::ArrayFormulaContext - missing XArrayFormulaTokens interface" );
    setBaseAddress( CellAddress( rArrayRange.Sheet, rArrayRange.StartColumn, rArrayRange.StartRow ), false );
}

void ArrayFormulaContext::setTokens( const ApiTokenSequence& rTokens )
{
    mxTokens->setArrayTokens( rTokens );
}

// Token sequence finalizer ===================================================

namespace {

class TokenSequenceFinalizer
{
public:
    explicit            TokenSequenceFinalizer(
                            const FunctionProvider& rFuncProv,
                            const ApiTokenSequence& rTokens );

    ApiTokenSequence    finalizeTokens();

private:
    typedef ::std::vector< const ApiToken* > ParameterPosVec;

    void                processTokens(
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd );

    const ApiToken*     processParameters(
                            const FunctionInfo& rFuncInfo,
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd );

    bool                isEmptyParameter(
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd ) const;

    OUString            getExternCallParameter(
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd ) const;

    const ApiToken*     skipParentheses(
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd ) const;

    const ApiToken*     findParameters(
                            ParameterPosVec& rParams,
                            const ApiToken* pToken,
                            const ApiToken* pTokenEnd ) const;

    void                appendCalcOnlyParameter(
                            const FunctionInfo& rFuncInfo,
                            size_t nParam );

    void                appendToken( const ApiToken& rToken );
    Any&                appendToken( sal_Int32 nOpCode );

private:
    typedef ::std::vector< ApiToken > ApiTokenVec;

    const FunctionProvider& mrFuncProv;
    const ApiTokenSequence& mrOldTokens;
    ApiTokenVec         maNewTokens;
};

// ----------------------------------------------------------------------------

TokenSequenceFinalizer::TokenSequenceFinalizer( const FunctionProvider& rFuncProv, const ApiTokenSequence& rTokens ) :
    mrFuncProv( rFuncProv ),
    mrOldTokens( rTokens )
{
}

ApiTokenSequence TokenSequenceFinalizer::finalizeTokens()
{
    sal_Int32 nLen = mrOldTokens.getLength();
    if( (nLen > 0) && maNewTokens.empty() )
    {
        maNewTokens.reserve( static_cast< size_t >( nLen + 16 ) );
        const ApiToken* pToken = mrOldTokens.getConstArray();
        processTokens( pToken, pToken + nLen );
    }
    return ContainerHelper::vectorToSequence( maNewTokens );
}

void TokenSequenceFinalizer::processTokens( const ApiToken* pToken, const ApiToken* pTokenEnd )
{
    while( pToken < pTokenEnd )
    {
        // push the current token into the vector
        appendToken( *pToken );
        // try to process a function, otherwise go to next token
        if( const FunctionInfo* pFuncInfo = mrFuncProv.getFuncInfoFromApiToken( *pToken ) )
            pToken = processParameters( *pFuncInfo, pToken + 1, pTokenEnd );
        else
            ++pToken;
    }
}

const ApiToken* TokenSequenceFinalizer::processParameters(
        const FunctionInfo& rFuncInfo, const ApiToken* pToken, const ApiToken* pTokenEnd )
{
    // remember position of the token containing the function op-code
    size_t nFuncNameIdx = maNewTokens.size() - 1;

    /*  Special handling for BIFF_FUNC_EXTERNCALL, used for add-ins in BIFF import. */
    bool bExternCall = rFuncInfo.mnBiffFuncId == BIFF_FUNC_EXTERNCALL;
    // replace OPCODE_EXTERNAL with OPCODE_NONAME to get #NAME! error in cell
    if( bExternCall )
        maNewTokens.back().OpCode = mrFuncProv.OPCODE_NONAME;

    // process a function, if an OPCODE_OPEN token is following
    OSL_ENSURE( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_OPEN), "TokenSequenceFinalizer::processParameters - OPCODE_OPEN expected" );
    if( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_OPEN) )
    {
        // append the OPCODE_OPEN token to the vector
        appendToken( mrFuncProv.OPCODE_OPEN );

        // store positions of OPCODE_OPEN, parameter separators, and OPCODE_CLOSE
        ParameterPosVec aParams;
        pToken = findParameters( aParams, pToken, pTokenEnd );
        OSL_ENSURE( aParams.size() >= 2, "TokenSequenceFinalizer::processParameters - missing tokens" );
        size_t nParamCount = aParams.size() - 1;

        if( (nParamCount == 1) && isEmptyParameter( aParams[ 0 ] + 1, aParams[ 1 ] ) )
        {
            /*  Empty pair of parentheses -> function call without parameters,
                process parameter, there might be spaces between parentheses. */
            processTokens( aParams[ 0 ] + 1, aParams[ 1 ] );
        }
        else
        {
            // process all parameters
            ParameterPosVec::const_iterator aPosIt = aParams.begin();
            FuncInfoParamClassIterator aClassIt( rFuncInfo );
            size_t nLastValidSize = maNewTokens.size();
            size_t nLastValidCount = 0;
            for( size_t nParam = 0; nParam < nParamCount; ++nParam, ++aPosIt, ++aClassIt )
            {
                // add embedded Calc-only parameters
                if( aClassIt.isCalcOnlyParam() )
                {
                    appendCalcOnlyParameter( rFuncInfo, nParam );
                    while( aClassIt.isCalcOnlyParam() ) ++aClassIt;
                }

                const ApiToken* pParamBegin = *aPosIt + 1;
                const ApiToken* pParamEnd = *(aPosIt + 1);
                bool bIsEmpty = isEmptyParameter( pParamBegin, pParamEnd );

                if( aClassIt.isExcelOnlyParam() )
                {
                    // process BIFF external calls, try to find a matching function
                    if( bExternCall && (nParam == 0) && !bIsEmpty )
                    {
                        OUString aName = getExternCallParameter( pParamBegin, pParamEnd );
                        if( const FunctionInfo* pExtFuncInfo = mrFuncProv.getFuncInfoFromExternCallName( aName ) )
                        {
                            maNewTokens[ nFuncNameIdx ].OpCode = pExtFuncInfo->mnApiOpCode;
                            // insert programmatic add-in function names
                            if( pExtFuncInfo->mnApiOpCode == mrFuncProv.OPCODE_EXTERNAL )
                                maNewTokens[ nFuncNameIdx ].Data <<= pExtFuncInfo->maExtProgName;
                        }
                    }
                }
                else
                {
                    // replace empty second and third parameter in IF function with zeros
                    if( (rFuncInfo.mnBiffFuncId == BIFF_FUNC_IF) && ((nParam == 1) || (nParam == 2)) && bIsEmpty )
                    {
                        appendToken( mrFuncProv.OPCODE_PUSH ) <<= static_cast< double >( 0.0 );
                        bIsEmpty = false;
                    }
                    else
                    {
                        // process all tokens of the parameter
                        processTokens( pParamBegin, pParamEnd );
                    }
                    // append parameter separator token
                    appendToken( mrFuncProv.OPCODE_SEP );
                }

                /*  #84453# Update size of new token sequence with valid parameters
                    to be able to remove trailing optional empty parameters. */
                if( !bIsEmpty || (nParam < rFuncInfo.mnMinParamCount) )
                {
                    nLastValidSize = maNewTokens.size();
                    nLastValidCount = nParam + 1;
                }
            }

            // #84453# remove trailing optional empty parameters
            maNewTokens.resize( nLastValidSize );

            // add trailing Calc-only parameters
            if( aClassIt.isCalcOnlyParam() )
                appendCalcOnlyParameter( rFuncInfo, nLastValidCount );

            // remove last parameter separator token
            if( maNewTokens.back().OpCode == mrFuncProv.OPCODE_SEP )
                maNewTokens.pop_back();
        }

        /*  Append the OPCODE_CLOSE token to the vector, but only if there is
            no OPCODE_BAD token at the end, this token already contains the
            trailing closing parentheses. */
        if( (pTokenEnd - 1)->OpCode != mrFuncProv.OPCODE_BAD )
            appendToken( mrFuncProv.OPCODE_CLOSE );
    }

    return pToken;
}

bool TokenSequenceFinalizer::isEmptyParameter( const ApiToken* pToken, const ApiToken* pTokenEnd ) const
{
    while( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_SPACES) ) ++pToken;
    if( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_MISSING) ) ++pToken;
    while( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_SPACES) ) ++pToken;
    return pToken == pTokenEnd;
}

OUString TokenSequenceFinalizer::getExternCallParameter( const ApiToken* pToken, const ApiToken* pTokenEnd ) const
{
    OUString aExtCallName;
    while( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_SPACES) ) ++pToken;
    if( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_MACRO) ) (pToken++)->Data >>= aExtCallName;
    while( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_SPACES) ) ++pToken;
    return (pToken == pTokenEnd) ? aExtCallName : OUString();
}

const ApiToken* TokenSequenceFinalizer::skipParentheses( const ApiToken* pToken, const ApiToken* pTokenEnd ) const
{
    // skip tokens between OPCODE_OPEN and OPCODE_CLOSE
    OSL_ENSURE( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_OPEN), "skipParentheses - OPCODE_OPEN expected" );
    ++pToken;
    while( (pToken < pTokenEnd) && (pToken->OpCode != mrFuncProv.OPCODE_CLOSE) )
    {
        if( pToken->OpCode == mrFuncProv.OPCODE_OPEN )
            pToken = skipParentheses( pToken, pTokenEnd );
        else
            ++pToken;
    }
    // skip the OPCODE_CLOSE token
    OSL_ENSURE( ((pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_CLOSE)) || ((pTokenEnd - 1)->OpCode == mrFuncProv.OPCODE_BAD), "skipParentheses - OPCODE_CLOSE expected" );
    return (pToken < pTokenEnd) ? (pToken + 1) : pTokenEnd;
}

const ApiToken* TokenSequenceFinalizer::findParameters( ParameterPosVec& rParams,
        const ApiToken* pToken, const ApiToken* pTokenEnd ) const
{
    // push position of OPCODE_OPEN
    OSL_ENSURE( (pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_OPEN), "TokenSequenceFinalizer::findParameters - OPCODE_OPEN expected" );
    rParams.push_back( pToken++ );

    // find positions of parameter separators
    while( (pToken < pTokenEnd) && (pToken->OpCode != mrFuncProv.OPCODE_CLOSE) )
    {
        if( pToken->OpCode == mrFuncProv.OPCODE_OPEN )
            pToken = skipParentheses( pToken, pTokenEnd );
        else if( pToken->OpCode == mrFuncProv.OPCODE_SEP )
            rParams.push_back( pToken++ );
        else
            ++pToken;
    }

    // push position of OPCODE_CLOSE
    OSL_ENSURE( ((pToken < pTokenEnd) && (pToken->OpCode == mrFuncProv.OPCODE_CLOSE)) || ((pTokenEnd - 1)->OpCode == mrFuncProv.OPCODE_BAD), "TokenSequenceFinalizer::findParameters - OPCODE_CLOSE expected" );
    rParams.push_back( pToken );
    return (pToken < pTokenEnd) ? (pToken + 1) : pTokenEnd;
}

void TokenSequenceFinalizer::appendCalcOnlyParameter( const FunctionInfo& rFuncInfo, size_t nParam )
{
    (void)nParam;   // prevent 'unused' warning
    switch( rFuncInfo.mnBiffFuncId )
    {
        case BIFF_FUNC_FLOOR:
        case BIFF_FUNC_CEILING:
            OSL_ENSURE( nParam == 2, "TokenSequenceFinalizer::appendCalcOnlyParameter - unexpected parameter index" );
            appendToken( mrFuncProv.OPCODE_PUSH ) <<= static_cast< double >( 1.0 );
            appendToken( mrFuncProv.OPCODE_SEP );
        break;
    }
}

void TokenSequenceFinalizer::appendToken( const ApiToken& rToken )
{
    if( (rToken.OpCode == mrFuncProv.OPCODE_MACRO) && !rToken.Data.hasValue() )
    {
        appendToken( mrFuncProv.OPCODE_ARRAY_OPEN );
        appendToken( mrFuncProv.OPCODE_PUSH ) <<= BiffHelper::calcDoubleFromError( BIFF_ERR_NAME );
        appendToken( mrFuncProv.OPCODE_ARRAY_CLOSE );
    }
    else
        maNewTokens.push_back( rToken );
}

Any& TokenSequenceFinalizer::appendToken( sal_Int32 nOpCode )
{
    maNewTokens.resize( maNewTokens.size() + 1 );
    maNewTokens.back().OpCode = nOpCode;
    return maNewTokens.back().Data;
}

} // namespace

// parser implementation base =================================================

class FormulaParserImpl : public GlobalDataHelper
{
public:
    explicit            FormulaParserImpl(
                            const GlobalDataHelper& rGlobalData,
                            const FunctionProvider& rFuncProv );

    /** Converts an XML formula string. */
    virtual void        importOoxFormula(
                            FormulaContext& rContext,
                            const OUString& rFormulaString );

    /** Imports and converts a BIFF token array from the passed stream. */
    virtual void        importBiffFormula(
                            FormulaContext& rContext,
                            BiffInputStream& rStrm, const sal_uInt16* pnFmlaSize );

    /** Finalizes the passed token array after import (e.g. adjusts function
        parameters) and sets the formula using the passed context. */
    void                setFormula(
                            FormulaContext& rContext,
                            const ApiTokenSequence& rTokens );

protected:
    /** Sets the current formula context used for importing a function. */
    inline void         setFormulaContext( FormulaContext& rContext ) { mpContext = &rContext; }
    /** Sets the current formula context used for importing a function. */
    inline FormulaContext& getFormulaContext() const { return *mpContext; }

    /** Finalizes the passed token array after import. */
    void                finalizeImport( const ApiTokenSequence& rTokens );

protected:
    const FunctionProvider& mrFuncProv;

private:
    FormulaContext*     mpContext;
};

// ----------------------------------------------------------------------------

FormulaParserImpl::FormulaParserImpl( const GlobalDataHelper& rGlobalData, const FunctionProvider& rFuncProv ) :
    GlobalDataHelper( rGlobalData ),
    mrFuncProv( rFuncProv ),
    mpContext( 0 )
{
}

void FormulaParserImpl::importOoxFormula( FormulaContext& /*rContext*/,
        const OUString& /*rFormulaString*/ )
{
    OSL_ENSURE( false, "FormulaParserImpl::importOoxFormula - not implemented" );
}

void FormulaParserImpl::importBiffFormula( FormulaContext& /*rContext*/,
        BiffInputStream& /*rStrm*/, const sal_uInt16* /*pnFmlaSize*/ )
{
    OSL_ENSURE( false, "FormulaParserImpl::importBiffFormula - not implemented" );
}

void FormulaParserImpl::setFormula( FormulaContext& rContext, const ApiTokenSequence& rTokens )
{
    setFormulaContext( rContext );
    finalizeImport( rTokens );
}

void FormulaParserImpl::finalizeImport( const ApiTokenSequence& rTokens )
{
    TokenSequenceFinalizer aFinalizer( mrFuncProv, rTokens );
    mpContext->setTokens( aFinalizer.finalizeTokens() );
}

// OOX parser implementation ==================================================

class OoxFormulaParserImpl : public FormulaParserImpl
{
public:
    explicit            OoxFormulaParserImpl(
                            const GlobalDataHelper& rGlobalData,
                            const FunctionProvider& rFuncProv );

    virtual void        importOoxFormula(
                            FormulaContext& rContext,
                            const OUString& rFormulaString );

private:
    Reference< XFormulaParser > mxParser;
    PropertySet         maParserProps;
    const OUString      maRefPosProp;
};

// ----------------------------------------------------------------------------

OoxFormulaParserImpl::OoxFormulaParserImpl( const GlobalDataHelper& rGlobalData, const FunctionProvider& rFuncProv ) :
    FormulaParserImpl( rGlobalData, rFuncProv ),
    maRefPosProp( CREATE_OUSTRING( "ReferencePosition" ) )
{
    try
    {
        Reference< XMultiServiceFactory > xFactory( getDocument(), UNO_QUERY_THROW );
        mxParser.set( xFactory->createInstance( CREATE_OUSTRING( "com.sun.star.sheet.FormulaParser" ) ), UNO_QUERY_THROW );
    }
    catch( Exception& )
    {
    }
    OSL_ENSURE( mxParser.is(), "OoxFormulaParserImpl::OoxFormulaParserImpl - cannot create formula parser" );
    maParserProps.set( mxParser );
    maParserProps.setProperty( CREATE_OUSTRING( "CompileEnglish" ), true );
    maParserProps.setProperty( CREATE_OUSTRING( "R1C1Notation" ), false );
    maParserProps.setProperty( CREATE_OUSTRING( "Compatibility3DNotation" ), true );
    maParserProps.setProperty( CREATE_OUSTRING( "IgnoreLeadingSpaces" ), false );
    maParserProps.setProperty( CREATE_OUSTRING( "OpCodeMap" ), mrFuncProv.getOoxParserMap() );
}

void OoxFormulaParserImpl::importOoxFormula(
        FormulaContext& rContext, const OUString& rFormulaString )
{
    if( mxParser.is() )
    {
        setFormulaContext( rContext );
        maParserProps.setProperty( maRefPosProp, getFormulaContext().getBaseAddress() );
        finalizeImport( mxParser->parseFormula( rFormulaString ) );
    }
}

// BIFF parser implementation =================================================

namespace {

/** A 2D formula cell reference struct with relative flags. */
struct BiffSingleRef2d
{
    sal_Int32           mnCol;              /// Column index.
    sal_Int32           mnRow;              /// Row index.
    bool                mbColRel;           /// True = relative column reference.
    bool                mbRowRel;           /// True = relative row reference.

    explicit            BiffSingleRef2d();

    void                readBiff2Data( BiffInputStream& rStrm, bool bRelativeAsOffset );
    void                readBiff8Data( BiffInputStream& rStrm, bool bRelativeAsOffset );

    void                setBiff2Data( sal_uInt8 nCol, sal_uInt16 nRow, bool bRelativeAsOffset );
    void                setBiff8Data( sal_uInt16 nCol, sal_uInt16 nRow, bool bRelativeAsOffset );
};

BiffSingleRef2d::BiffSingleRef2d() :
    mnCol( 0 ),
    mnRow( 0 ),
    mbColRel( false ),
    mbRowRel( false )
{
}

void BiffSingleRef2d::readBiff2Data( BiffInputStream& rStrm, bool bRelativeAsOffset )
{
    sal_uInt16 nRow;
    sal_uInt8 nCol;
    rStrm >> nRow >> nCol;
    setBiff2Data( nCol, nRow, bRelativeAsOffset );
}

void BiffSingleRef2d::readBiff8Data( BiffInputStream& rStrm, bool bRelativeAsOffset )
{
    sal_uInt16 nRow, nCol;
    rStrm >> nRow >> nCol;
    setBiff8Data( nCol, nRow, bRelativeAsOffset );
}

void BiffSingleRef2d::setBiff2Data( sal_uInt8 nCol, sal_uInt16 nRow, bool bRelativeAsOffset )
{
    mnCol = nCol;
    mnRow = nRow & BIFF_TOK_REF_ROWMASK;
    mbColRel = getFlag( nRow, BIFF_TOK_REF_COLREL );
    mbRowRel = getFlag( nRow, BIFF_TOK_REF_ROWREL );
    if( bRelativeAsOffset && mbColRel && (mnCol >= 0x80) )
        mnCol -= 0x100;
    if( bRelativeAsOffset && mbRowRel && (mnRow > (BIFF_TOK_REF_ROWMASK >> 1)) )
        mnRow -= (BIFF_TOK_REF_ROWMASK + 1);
}

void BiffSingleRef2d::setBiff8Data( sal_uInt16 nCol, sal_uInt16 nRow, bool bRelativeAsOffset )
{
    mnCol = nCol & BIFF_TOK_REF_COLMASK;
    mnRow = nRow;
    mbColRel = getFlag( nCol, BIFF_TOK_REF_COLREL );
    mbRowRel = getFlag( nCol, BIFF_TOK_REF_ROWREL );
    if( bRelativeAsOffset && mbColRel && (mnCol > (BIFF_TOK_REF_COLMASK >> 1)) )
        mnCol -= (BIFF_TOK_REF_COLMASK + 1);
    if( bRelativeAsOffset && mbRowRel && (mnRow >= 0x8000) )
        mnRow -= 0x10000;
}

// ----------------------------------------------------------------------------

/** A 2D formula cell range reference struct with relative flags. */
struct BiffComplexRef2d
{
    BiffSingleRef2d     maRef1;             /// Start (top-left) cell address.
    BiffSingleRef2d     maRef2;             /// End (bottom-right) cell address.

    void                readBiff2Data( BiffInputStream& rStrm, bool bRelativeAsOffset );
    void                readBiff8Data( BiffInputStream& rStrm, bool bRelativeAsOffset );
};

void BiffComplexRef2d::readBiff2Data( BiffInputStream& rStrm, bool bRelativeAsOffset )
{
    sal_uInt16 nRow1, nRow2;
    sal_uInt8 nCol1, nCol2;
    rStrm >> nRow1 >> nRow2 >> nCol1 >> nCol2;
    maRef1.setBiff2Data( nCol1, nRow1, bRelativeAsOffset );
    maRef2.setBiff2Data( nCol2, nRow2, bRelativeAsOffset );
}

void BiffComplexRef2d::readBiff8Data( BiffInputStream& rStrm, bool bRelativeAsOffset )
{
    sal_uInt16 nRow1, nRow2, nCol1, nCol2;
    rStrm >> nRow1 >> nRow2 >> nCol1 >> nCol2;
    maRef1.setBiff8Data( nCol1, nRow1, bRelativeAsOffset );
    maRef2.setBiff8Data( nCol2, nRow2, bRelativeAsOffset );
}

// ----------------------------------------------------------------------------

/** A natural language reference struct with relative flag. */
struct BiffNlr
{
    sal_uInt16          mnCol;              /// Column index.
    sal_uInt16          mnRow;              /// Row index.
    bool                mbRel;              /// True = relative column/row reference.

    explicit            BiffNlr();

    void                readBiff8Data( BiffInputStream& rStrm );
};

BiffNlr::BiffNlr() :
    mnCol( 0 ),
    mnRow( 0 ),
    mbRel( false )
{
}

void BiffNlr::readBiff8Data( BiffInputStream& rStrm )
{
    rStrm >> mnRow >> mnCol;
    mbRel = getFlag( mnCol, BIFF_TOK_NLR_REL );
    mnCol &= BIFF_TOK_NLR_MASK;
}

bool lclIsValidNlrStack( const BiffAddress& rBiffAddr1, const BiffAddress& rBiffAddr2, bool bRow )
{
    return bRow ?
        ((rBiffAddr1.mnRow == rBiffAddr2.mnRow) && (rBiffAddr1.mnCol + 1 == rBiffAddr2.mnCol)) :
        ((rBiffAddr1.mnCol == rBiffAddr2.mnCol) && (rBiffAddr1.mnRow + 1 == rBiffAddr2.mnRow));
}

bool lclIsValidNlrRange( const BiffNlr& rBiffNlr, const BiffRange& rBiffRange, bool bRow )
{
    return bRow ?
        ((rBiffNlr.mnRow == rBiffRange.maFirst.mnRow) && (rBiffNlr.mnCol + 1 == rBiffRange.maFirst.mnCol) && (rBiffRange.maFirst.mnRow == rBiffRange.maLast.mnRow)) :
        ((rBiffNlr.mnCol == rBiffRange.maFirst.mnCol) && (rBiffNlr.mnRow + 1 == rBiffRange.maFirst.mnRow) && (rBiffRange.maFirst.mnCol == rBiffRange.maLast.mnCol));
}

} // namespace

// ----------------------------------------------------------------------------

class BiffFormulaParserImpl : public FormulaParserImpl
{
public:
    explicit            BiffFormulaParserImpl(
                            const GlobalDataHelper& rGlobalData,
                            const FunctionProvider& rFuncProv );

    virtual void        importBiffFormula(
                            FormulaContext& rContext,
                            BiffInputStream& rStrm, const sal_uInt16* pnFmlaSize );

private:
    // import token contents and create API formula token ---------------------

    bool                importTokenNotAvailable( BiffInputStream& rStrm );
    bool                importRefTokenNotAvailable( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importStrToken2( BiffInputStream& rStrm );
    bool                importStrToken8( BiffInputStream& rStrm );
    bool                importAttrToken( BiffInputStream& rStrm );
    bool                importSpaceToken3( BiffInputStream& rStrm );
    bool                importSpaceToken4( BiffInputStream& rStrm );
    bool                importSheetToken2( BiffInputStream& rStrm );
    bool                importSheetToken3( BiffInputStream& rStrm );
    bool                importEndSheetToken2( BiffInputStream& rStrm );
    bool                importEndSheetToken3( BiffInputStream& rStrm );
    bool                importNlrToken( BiffInputStream& rStrm );
    bool                importArrayToken( BiffInputStream& rStrm );
    bool                importNameToken( BiffInputStream& rStrm );
    bool                importRefToken2( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importRefToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importAreaToken2( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importAreaToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importRef3dToken5( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importRef3dToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importArea3dToken5( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importArea3dToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset );
    bool                importMemAreaToken( BiffInputStream& rStrm, bool bAddData );
    bool                importMemFuncToken( BiffInputStream& rStrm );
    bool                importNameXToken( BiffInputStream& rStrm );
    bool                importFuncToken2( BiffInputStream& rStrm );
    bool                importFuncToken4( BiffInputStream& rStrm );
    bool                importFuncVarToken2( BiffInputStream& rStrm );
    bool                importFuncVarToken4( BiffInputStream& rStrm );
    bool                importFuncCEToken( BiffInputStream& rStrm );
    bool                importExpToken2( BiffInputStream& rStrm );
    bool                importExpToken3( BiffInputStream& rStrm );
    bool                importTblToken2( BiffInputStream& rStrm );
    bool                importTblToken3( BiffInputStream& rStrm );

    bool                importNlrAddrToken( BiffInputStream& rStrm, bool bRow );
    bool                importNlrRangeToken( BiffInputStream& rStrm );
    bool                importNlrSAddrToken( BiffInputStream& rStrm, bool bRow );
    bool                importNlrSRangeToken( BiffInputStream& rStrm );
    bool                importNlrErrToken( BiffInputStream& rStrm, sal_uInt16 nIgnore );

    sal_Int32           readRefId( BiffInputStream& rStrm );
    sal_uInt16          readNameId( BiffInputStream& rStrm );
    LinkSheetRange      readSheetRange5( BiffInputStream& rStrm );
    LinkSheetRange      readSheetRange8( BiffInputStream& rStrm );

    void                swapStreamPosition( BiffInputStream& rStrm );
    void                ignoreMemAreaAddData( BiffInputStream& rStrm );
    bool                readNlrSAddrAddData( BiffNlr& orBiffNlr, BiffInputStream& rStrm, bool bRow );
    bool                readNlrSRangeAddData( BiffNlr& orBiffNlr, bool& orbIsRow, BiffInputStream& rStrm );

    // push API operand or operator -------------------------------------------

    bool                pushOperandToken( sal_Int32 nOpCode, sal_Int32 nSpaces );
    template< typename Type >
    bool                pushValueOperandToken( const Type& rValue, sal_Int32 nOpCode, sal_Int32 nSpaces );
    bool                pushParenthesesOperandToken( sal_Int32 nOpeningSpaces, sal_Int32 nClosingSpaces );
    bool                pushUnaryPreOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces );
    bool                pushUnaryPostOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces );
    bool                pushBinaryOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces );
    bool                pushParenthesesOperatorToken( sal_Int32 nOpeningSpaces, sal_Int32 nClosingSpaces );

    bool                pushOperand( sal_Int32 nOpCode );
    template< typename Type >
    bool                pushValueOperand( const Type& rValue, sal_Int32 nOpCode );
    template< typename Type >
    inline bool         pushValueOperand( const Type& rValue ) { return pushValueOperand( rValue, mrFuncProv.OPCODE_PUSH ); }
    bool                pushParenthesesOperand();
    bool                pushDelAddressOperand();
    bool                pushUnaryPreOperator( sal_Int32 nOpCode );
    bool                pushUnaryPostOperator( sal_Int32 nOpCode );
    bool                pushBinaryOperator( sal_Int32 nOpCode );
    bool                pushRangeOperator();
    bool                pushParenthesesOperator();
    bool                pushFunctionOperator( sal_Int32 nOpCode, size_t nParamCount );

    size_t              getOperandSize( size_t nOpCountFromEnd, size_t nOpIndex );
    void                pushOperandSize( size_t nSize );
    size_t              popOperandSize();

    ApiToken&           getOperandToken( size_t nOpCountFromEnd, size_t nOpIndex );
    void                removeOperand( size_t nOpCountFromEnd, size_t nOpIndex );
    void                removeLastOperands( size_t nOpCountFromEnd );

    Any&                appendRawToken( sal_Int32 nOpCode );
    Any&                insertRawToken( sal_Int32 nOpCode, size_t nIndexFromEnd );

    size_t              appendSpacesToken( sal_Int32 nSpaces );
    size_t              insertSpacesToken( sal_Int32 nSpaces, size_t nIndexFromEnd );
    bool                resetSpaces();

    // convert BIFF token and push API operand or operator --------------------

    bool                pushBiffBool( sal_uInt8 nValue );
    bool                pushBiffError( sal_uInt8 nErrorCode );
    bool                pushBiffSingleRef2d( const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset );
    bool                pushBiffComplexRef2d( const BiffComplexRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset );
    bool                pushBiffSingleRef3d( const LinkSheetRange& rSheetRange, const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset );
    bool                pushBiffComplexRef3d( const LinkSheetRange& rSheetRange, const BiffComplexRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset );
    bool                pushBiffNlrAddr( const BiffNlr& rBiffNlr, bool bRow );
    bool                pushBiffNlrRange( const BiffNlr& rBiffNlr, const BiffRange& rRange );
    bool                pushBiffNlrSAddr( const BiffNlr& rBiffNlr, bool bRow );
    bool                pushBiffNlrSRange( const BiffNlr& rBiffNlr, const BiffRange& rRange, bool bRow );
    bool                pushBiffName( sal_uInt16 nNameId );
    bool                pushBiffExtName( sal_Int32 nRefId, sal_uInt16 nNameId );
    bool                pushBiffAnalysisName( sal_Int32 nRefId, sal_uInt16 nNameId );
    bool                pushBiffFunction( const FunctionInfo& rFuncInfo, sal_uInt8 nParamCount );
    bool                pushBiffFunction( sal_uInt16 nFuncId );
    bool                pushBiffFunction( sal_uInt16 nFuncId, sal_uInt8 nParamCount );

    void                initializeRef2d( SingleReference& orApiRef ) const;
    void                initializeRef3d( SingleReference& orApiRef, sal_Int32 nSheet ) const;
    void                convertColRow( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, bool bRelativeAsOffset ) const;
    void                convertColRowDel( SingleReference& orApiRef ) const;
    void                convertSingleRef2d( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset ) const;
    void                convertSingleRef3d( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, sal_Int32 nSheet, bool bDeleted, bool bRelativeAsOffset ) const;

    bool                processExpToken( const BiffAddress& rBaseAddr );
    bool                processTblToken( const BiffAddress& rBaseAddr );

    // ------------------------------------------------------------------------
private:
    typedef ::std::vector< ApiToken >   ApiTokenVec;
    typedef ::std::vector< size_t >     SizeTypeVec;
    typedef bool (BiffFormulaParserImpl::*ImportTokenFunc)( BiffInputStream& );
    typedef bool (BiffFormulaParserImpl::*ImportRefTokenFunc)( BiffInputStream&, bool, bool );

    const sal_uInt16    mnMaxCol;
    const sal_uInt16    mnMaxRow;
    ApiTokenVec         maTokenStorage;
    SizeTypeVec         maTokenIndexes;
    SizeTypeVec         maOperandSizeStack;
    ImportTokenFunc     mpImportStrToken;           /// Pointer to tStr import function (string constant).
    ImportTokenFunc     mpImportSpaceToken;         /// Pointer to tAttrSpace import function (spaces/line breaks).
    ImportTokenFunc     mpImportSheetToken;         /// Pointer to tSheet import function (external reference).
    ImportTokenFunc     mpImportEndSheetToken;      /// Pointer to tEndSheet import function (end of external reference).
    ImportTokenFunc     mpImportNlrToken;           /// Pointer to tNlr import function (natural language reference).
    ImportRefTokenFunc  mpImportRefToken;           /// Pointer to tRef import function (2d cell reference).
    ImportRefTokenFunc  mpImportAreaToken;          /// Pointer to tArea import function (2d area reference).
    ImportRefTokenFunc  mpImportRef3dToken;         /// Pointer to tRef3d import function (3d cell reference).
    ImportRefTokenFunc  mpImportArea3dToken;        /// Pointer to tArea3d import function (3d area reference).
    ImportTokenFunc     mpImportNameXToken;         /// Pointer to tNameX import function (external name).
    ImportTokenFunc     mpImportFuncToken;          /// Pointer to tFunc import function (function with fixed parameter count).
    ImportTokenFunc     mpImportFuncVarToken;       /// Pointer to tFuncVar import function (function with variable parameter count).
    ImportTokenFunc     mpImportFuncCEToken;        /// Pointer to tFuncCE import function (command macro call).
    ImportTokenFunc     mpImportExpToken;           /// Pointer to tExp import function (array/shared formula).
    ImportTokenFunc     mpImportTblToken;           /// Pointer to tTbl import function (table operation).
    sal_uInt32          mnAddDataPos;               /// Current stream position for additional data (tArray, tMemArea, tNlr).
    sal_Int32           mnLeadingSpaces;            /// Current number of spaces before next token.
    sal_Int32           mnOpeningSpaces;            /// Current number of spaces before opening parenthesis.
    sal_Int32           mnClosingSpaces;            /// Current number of spaces before closing parenthesis.
    sal_Int32           mnCurrRefId;                /// Current ref-id from tSheet token (BIFF2-BIFF4 only).
    sal_uInt16          mnAttrDataSize;             /// Size of one tAttr data element.
    sal_uInt16          mnArraySize;                /// Size of tArray data.
    sal_uInt16          mnNameSize;                 /// Size of tName data.
    sal_uInt16          mnMemAreaSize;              /// Size of tMemArea data.
    sal_uInt16          mnMemFuncSize;              /// Size of tMemFunc data.
    sal_uInt16          mnRefIdSize;                /// Size of unused data following a reference identifier.
};

// ----------------------------------------------------------------------------

BiffFormulaParserImpl::BiffFormulaParserImpl( const GlobalDataHelper& rGlobalData, const FunctionProvider& rFuncProv ) :
    FormulaParserImpl( rGlobalData, rFuncProv ),
    mnMaxCol( static_cast< sal_uInt16 >( rGlobalData.getAddressConverter().getMaxAddress().Column ) ),
    mnMaxRow( static_cast< sal_uInt16 >( rGlobalData.getAddressConverter().getMaxAddress().Row ) ),
    mnAddDataPos( 0 ),
    mnLeadingSpaces( 0 ),
    mnOpeningSpaces( 0 ),
    mnClosingSpaces( 0 ),
    mnCurrRefId( 0 )
{
    maTokenStorage.reserve( BIFF_TOKARR_MAXLEN );
    maTokenIndexes.reserve( BIFF_TOKARR_MAXLEN );
    maOperandSizeStack.reserve( 256 );

    switch( getBiff() )
    {
        case BIFF2:
            mpImportStrToken = &BiffFormulaParserImpl::importStrToken2;
            mpImportSpaceToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportSheetToken = &BiffFormulaParserImpl::importSheetToken2;
            mpImportEndSheetToken = &BiffFormulaParserImpl::importEndSheetToken2;
            mpImportNlrToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportRefToken = &BiffFormulaParserImpl::importRefToken2;
            mpImportAreaToken = &BiffFormulaParserImpl::importAreaToken2;
            mpImportRef3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportArea3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportNameXToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportFuncToken = &BiffFormulaParserImpl::importFuncToken2;
            mpImportFuncVarToken = &BiffFormulaParserImpl::importFuncVarToken2;
            mpImportFuncCEToken = &BiffFormulaParserImpl::importFuncCEToken;
            mpImportExpToken = &BiffFormulaParserImpl::importExpToken2;
            mpImportTblToken = &BiffFormulaParserImpl::importTblToken2;
            mnAttrDataSize = 1;
            mnArraySize = 6;
            mnNameSize = 5;
            mnMemAreaSize = 4;
            mnMemFuncSize = 1;
            mnRefIdSize = 1;
        break;
        case BIFF3:
            mpImportStrToken = &BiffFormulaParserImpl::importStrToken2;
            mpImportSpaceToken = &BiffFormulaParserImpl::importSpaceToken3;
            mpImportSheetToken = &BiffFormulaParserImpl::importSheetToken3;
            mpImportEndSheetToken = &BiffFormulaParserImpl::importEndSheetToken3;
            mpImportNlrToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportRefToken = &BiffFormulaParserImpl::importRefToken2;
            mpImportAreaToken = &BiffFormulaParserImpl::importAreaToken2;
            mpImportRef3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportArea3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportNameXToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportFuncToken = &BiffFormulaParserImpl::importFuncToken2;
            mpImportFuncVarToken = &BiffFormulaParserImpl::importFuncVarToken2;
            mpImportFuncCEToken = &BiffFormulaParserImpl::importFuncCEToken;
            mpImportExpToken = &BiffFormulaParserImpl::importExpToken3;
            mpImportTblToken = &BiffFormulaParserImpl::importTblToken3;
            mnAttrDataSize = 2;
            mnArraySize = 7;
            mnNameSize = 8;
            mnMemAreaSize = 6;
            mnMemFuncSize = 2;
            mnRefIdSize = 2;
        break;
        case BIFF4:
            mpImportStrToken = &BiffFormulaParserImpl::importStrToken2;
            mpImportSpaceToken = &BiffFormulaParserImpl::importSpaceToken4;
            mpImportSheetToken = &BiffFormulaParserImpl::importSheetToken3;
            mpImportEndSheetToken = &BiffFormulaParserImpl::importEndSheetToken3;
            mpImportNlrToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportRefToken = &BiffFormulaParserImpl::importRefToken2;
            mpImportAreaToken = &BiffFormulaParserImpl::importAreaToken2;
            mpImportRef3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportArea3dToken = &BiffFormulaParserImpl::importRefTokenNotAvailable;
            mpImportNameXToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportFuncToken = &BiffFormulaParserImpl::importFuncToken4;
            mpImportFuncVarToken = &BiffFormulaParserImpl::importFuncVarToken4;
            mpImportFuncCEToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportExpToken = &BiffFormulaParserImpl::importExpToken3;
            mpImportTblToken = &BiffFormulaParserImpl::importTblToken3;
            mnAttrDataSize = 2;
            mnArraySize = 7;
            mnNameSize = 8;
            mnMemAreaSize = 6;
            mnMemFuncSize = 2;
            mnRefIdSize = 2;
        break;
        case BIFF5:
            mpImportStrToken = &BiffFormulaParserImpl::importStrToken2;
            mpImportSpaceToken = &BiffFormulaParserImpl::importSpaceToken4;
            mpImportSheetToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportEndSheetToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportNlrToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportRefToken = &BiffFormulaParserImpl::importRefToken2;
            mpImportAreaToken = &BiffFormulaParserImpl::importAreaToken2;
            mpImportRef3dToken = &BiffFormulaParserImpl::importRef3dToken5;
            mpImportArea3dToken = &BiffFormulaParserImpl::importArea3dToken5;
            mpImportNameXToken = &BiffFormulaParserImpl::importNameXToken;
            mpImportFuncToken = &BiffFormulaParserImpl::importFuncToken4;
            mpImportFuncVarToken = &BiffFormulaParserImpl::importFuncVarToken4;
            mpImportFuncCEToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportExpToken = &BiffFormulaParserImpl::importExpToken3;
            mpImportTblToken = &BiffFormulaParserImpl::importTblToken3;
            mnAttrDataSize = 2;
            mnArraySize = 7;
            mnNameSize = 12;
            mnMemAreaSize = 6;
            mnMemFuncSize = 2;
            mnRefIdSize = 8;
        break;
        case BIFF8:
            mpImportStrToken = &BiffFormulaParserImpl::importStrToken8;
            mpImportSpaceToken = &BiffFormulaParserImpl::importSpaceToken4;
            mpImportSheetToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportEndSheetToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportNlrToken = &BiffFormulaParserImpl::importNlrToken;
            mpImportRefToken = &BiffFormulaParserImpl::importRefToken8;
            mpImportAreaToken = &BiffFormulaParserImpl::importAreaToken8;
            mpImportRef3dToken = &BiffFormulaParserImpl::importRef3dToken8;
            mpImportArea3dToken = &BiffFormulaParserImpl::importArea3dToken8;
            mpImportNameXToken = &BiffFormulaParserImpl::importNameXToken;
            mpImportFuncToken = &BiffFormulaParserImpl::importFuncToken4;
            mpImportFuncVarToken = &BiffFormulaParserImpl::importFuncVarToken4;
            mpImportFuncCEToken = &BiffFormulaParserImpl::importTokenNotAvailable;
            mpImportExpToken = &BiffFormulaParserImpl::importExpToken3;
            mpImportTblToken = &BiffFormulaParserImpl::importTblToken3;
            mnAttrDataSize = 2;
            mnArraySize = 7;
            mnNameSize = 2;
            mnMemAreaSize = 6;
            mnMemFuncSize = 2;
            mnRefIdSize = 0;
        break;
        case BIFF_UNKNOWN: break;
    }
}

void BiffFormulaParserImpl::importBiffFormula( FormulaContext& rContext,
        BiffInputStream& rStrm, const sal_uInt16* pnFmlaSize )
{
    setFormulaContext( rContext );
    maTokenStorage.clear();
    maTokenIndexes.clear();
    maOperandSizeStack.clear();
    resetSpaces();
    mnCurrRefId = 0;

    sal_uInt16 nFmlaSize = pnFmlaSize ? *pnFmlaSize : ((getBiff() == BIFF2) ? rStrm.readuInt8() : rStrm.readuInt16());
    sal_uInt32 nEndPos = mnAddDataPos = rStrm.getRecPos() + nFmlaSize;
    bool bRelativeAsOffset = getFormulaContext().isRelativeAsOffset();

    bool bOk = true;
    while( bOk && rStrm.isValid() && (rStrm.getRecPos() < nEndPos) )
    {
        sal_uInt8 nTokenId;
        rStrm >> nTokenId;
        sal_uInt8 nTokenClass = nTokenId & BIFF_TOKCLASS_MASK;
        sal_uInt8 nBaseId = nTokenId & BIFF_TOKID_MASK;

        if( nTokenClass == BIFF_TOKCLASS_NONE )
        {
            // base tokens
            switch( nBaseId )
            {
                case BIFF_TOKID_EXP:        bOk = (this->*mpImportExpToken)( rStrm );                   break;
                case BIFF_TOKID_TBL:        bOk = (this->*mpImportTblToken)( rStrm );                   break;
                case BIFF_TOKID_ADD:        bOk = pushBinaryOperator( mrFuncProv.OPCODE_ADD );          break;
                case BIFF_TOKID_SUB:        bOk = pushBinaryOperator( mrFuncProv.OPCODE_SUB );          break;
                case BIFF_TOKID_MUL:        bOk = pushBinaryOperator( mrFuncProv.OPCODE_MULT );         break;
                case BIFF_TOKID_DIV:        bOk = pushBinaryOperator( mrFuncProv.OPCODE_DIV );          break;
                case BIFF_TOKID_POWER:      bOk = pushBinaryOperator( mrFuncProv.OPCODE_POWER );        break;
                case BIFF_TOKID_CONCAT:     bOk = pushBinaryOperator( mrFuncProv.OPCODE_CONCAT );       break;
                case BIFF_TOKID_LT:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_LESS );         break;
                case BIFF_TOKID_LE:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_LESS_EQUAL );   break;
                case BIFF_TOKID_EQ:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_EQUAL );        break;
                case BIFF_TOKID_GE:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_GREATER_EQUAL );break;
                case BIFF_TOKID_GT:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_GREATER );      break;
                case BIFF_TOKID_NE:         bOk = pushBinaryOperator( mrFuncProv.OPCODE_NOT_EQUAL );    break;
                case BIFF_TOKID_ISECT:      bOk = pushBinaryOperator( mrFuncProv.OPCODE_INTERSECT );    break;
                case BIFF_TOKID_LIST:       bOk = pushBinaryOperator( mrFuncProv.OPCODE_LIST );         break;
                case BIFF_TOKID_RANGE:      bOk = pushRangeOperator();                                  break;  // #i48496# needs extra processing
                case BIFF_TOKID_UPLUS:      bOk = pushUnaryPreOperator( mrFuncProv.OPCODE_PLUS_SIGN );  break;
                case BIFF_TOKID_UMINUS:     bOk = pushUnaryPreOperator( mrFuncProv.OPCODE_MINUS_SIGN ); break;
                case BIFF_TOKID_PERCENT:    bOk = pushUnaryPostOperator( mrFuncProv.OPCODE_PERCENT );   break;
                case BIFF_TOKID_PAREN:      bOk = pushParenthesesOperator();                            break;
                case BIFF_TOKID_MISSARG:    bOk = pushOperand( mrFuncProv.OPCODE_MISSING );             break;
                case BIFF_TOKID_STR:        bOk = (this->*mpImportStrToken)( rStrm );                   break;
                case BIFF_TOKID_NLR:        bOk = (this->*mpImportNlrToken)( rStrm );                   break;
                case BIFF_TOKID_ATTR:       bOk = importAttrToken( rStrm );                             break;
                case BIFF_TOKID_SHEET:      bOk = (this->*mpImportSheetToken)( rStrm );                 break;
                case BIFF_TOKID_ENDSHEET:   bOk = (this->*mpImportEndSheetToken)( rStrm );              break;
                case BIFF_TOKID_ERR:        bOk = pushBiffError( rStrm.readuInt8() );                   break;
                case BIFF_TOKID_BOOL:       bOk = pushBiffBool( rStrm.readuInt8() );                    break;
                case BIFF_TOKID_INT:        bOk = pushValueOperand< double >( rStrm.readuInt16() );     break;
                case BIFF_TOKID_NUM:        bOk = pushValueOperand( rStrm.readDouble() );               break;
                default: bOk = false;
            }
        }
        else
        {
            // classified tokens
            switch( nBaseId )
            {
                case BIFF_TOKID_ARRAY:      bOk = importArrayToken( rStrm );                                        break;
                case BIFF_TOKID_FUNC:       bOk = (this->*mpImportFuncToken)( rStrm );                              break;
                case BIFF_TOKID_FUNCVAR:    bOk = (this->*mpImportFuncVarToken)( rStrm );                           break;
                case BIFF_TOKID_NAME:       bOk = importNameToken( rStrm );                                         break;
                case BIFF_TOKID_REF:        bOk = (this->*mpImportRefToken)( rStrm, false, false );                 break;
                case BIFF_TOKID_AREA:       bOk = (this->*mpImportAreaToken)( rStrm, false, false );                break;
                case BIFF_TOKID_MEMAREA:    bOk = importMemAreaToken( rStrm, true );                                break;
                case BIFF_TOKID_MEMERR:     bOk = importMemAreaToken( rStrm, false );                               break;
                case BIFF_TOKID_MEMNOMEM:   bOk = importMemAreaToken( rStrm, false );                               break;
                case BIFF_TOKID_MEMFUNC:    bOk = importMemFuncToken( rStrm );                                      break;
                case BIFF_TOKID_REFERR:     bOk = (this->*mpImportRefToken)( rStrm, true, false );                  break;
                case BIFF_TOKID_AREAERR:    bOk = (this->*mpImportAreaToken)( rStrm, true, false );                 break;
                case BIFF_TOKID_REFN:       bOk = (this->*mpImportRefToken)( rStrm, false, true );                  break;
                case BIFF_TOKID_AREAN:      bOk = (this->*mpImportAreaToken)( rStrm, false, true );                 break;
                case BIFF_TOKID_MEMAREAN:   bOk = importMemFuncToken( rStrm );                                      break;
                case BIFF_TOKID_MEMNOMEMN:  bOk = importMemFuncToken( rStrm );                                      break;
                case BIFF_TOKID_FUNCCE:     bOk = (this->*mpImportFuncCEToken)( rStrm );                            break;
                case BIFF_TOKID_NAMEX:      bOk = (this->*mpImportNameXToken)( rStrm );                             break;
                case BIFF_TOKID_REF3D:      bOk = (this->*mpImportRef3dToken)( rStrm, false, bRelativeAsOffset );   break;
                case BIFF_TOKID_AREA3D:     bOk = (this->*mpImportArea3dToken)( rStrm, false, bRelativeAsOffset );  break;
                case BIFF_TOKID_REFERR3D:   bOk = (this->*mpImportRef3dToken)( rStrm, true, bRelativeAsOffset );    break;
                case BIFF_TOKID_AREAERR3D:  bOk = (this->*mpImportArea3dToken)( rStrm, true, bRelativeAsOffset );   break;
                default: bOk = false;
            }
        }
    }
    if( bOk )
        bOk = rStrm.getRecPos() == nEndPos;

    // build and finalize the token sequence
    if( bOk )
    {
        ApiTokenSequence aTokens( static_cast< sal_Int32 >( maTokenIndexes.size() ) );
        ApiToken* pToken = aTokens.getArray();
        for( SizeTypeVec::const_iterator aIt = maTokenIndexes.begin(), aEnd = maTokenIndexes.end(); aIt != aEnd; ++aIt, ++pToken )
            *pToken = maTokenStorage[ *aIt ];
        finalizeImport( aTokens );
    }

    // seek behind additional token data of tArray, tMemArea, tNlr tokens
    rStrm.seek( mnAddDataPos );
}

// import token contents and create API formula token -------------------------

bool BiffFormulaParserImpl::importTokenNotAvailable( BiffInputStream& )
{
    // dummy function for pointer-to-member-function
    return false;
}

bool BiffFormulaParserImpl::importRefTokenNotAvailable( BiffInputStream&, bool, bool )
{
    // dummy function for pointer-to-member-function
    return false;
}

bool BiffFormulaParserImpl::importStrToken2( BiffInputStream& rStrm )
{
    return pushValueOperand( rStrm.readByteString( false, getTextEncoding() ) );
}

bool BiffFormulaParserImpl::importStrToken8( BiffInputStream& rStrm )
{
    // read flags field for empty strings also
    return pushValueOperand( rStrm.readUniString( rStrm.readuInt8() ) );
}

bool BiffFormulaParserImpl::importAttrToken( BiffInputStream& rStrm )
{
    bool bOk = true;
    sal_uInt8 nType;
    rStrm >> nType;
    switch( nType )
    {
        case BIFF_TOK_ATTR_VOLATILE:
        case BIFF_TOK_ATTR_IF:
        case BIFF_TOK_ATTR_SKIP:
        case BIFF_TOK_ATTR_ASSIGN:
            rStrm.ignore( mnAttrDataSize );
        break;
        case BIFF_TOK_ATTR_CHOOSE:
            rStrm.ignore( mnAttrDataSize * (1 + ((getBiff() == BIFF2) ? rStrm.readuInt8() : rStrm.readuInt16())) );
        break;
        case BIFF_TOK_ATTR_SUM:
            rStrm.ignore( mnAttrDataSize );
            bOk = pushBiffFunction( BIFF_FUNC_SUM, 1 );
        break;
        case BIFF_TOK_ATTR_SPACE:
        case BIFF_TOK_ATTR_SPACE_VOLATILE:
            bOk = (this->*mpImportSpaceToken)( rStrm );
        break;
        default:
            bOk = false;
    }
    return bOk;
}

bool BiffFormulaParserImpl::importSpaceToken3( BiffInputStream& rStrm )
{
    rStrm.ignore( 2 );
    return true;
}

bool BiffFormulaParserImpl::importSpaceToken4( BiffInputStream& rStrm )
{
    sal_uInt8 nType, nCount;
    rStrm >> nType >> nCount;
    switch( nType )
    {
        case BIFF_TOK_ATTR_SPACE_SP:
        case BIFF_TOK_ATTR_SPACE_BR:
            mnLeadingSpaces += nCount;
        break;
        case BIFF_TOK_ATTR_SPACE_SP_OPEN:
        case BIFF_TOK_ATTR_SPACE_BR_OPEN:
            mnOpeningSpaces += nCount;
        break;
        case BIFF_TOK_ATTR_SPACE_SP_CLOSE:
        case BIFF_TOK_ATTR_SPACE_BR_CLOSE:
            mnClosingSpaces += nCount;
        break;
    }
    return true;
}

bool BiffFormulaParserImpl::importSheetToken2( BiffInputStream& rStrm )
{
    rStrm.ignore( 4 );
    mnCurrRefId = readRefId( rStrm );
    return true;
}

bool BiffFormulaParserImpl::importSheetToken3( BiffInputStream& rStrm )
{
    rStrm.ignore( 6 );
    mnCurrRefId = readRefId( rStrm );
    return true;
}

bool BiffFormulaParserImpl::importEndSheetToken2( BiffInputStream& rStrm )
{
    rStrm.ignore( 3 );
    mnCurrRefId = 0;
    return true;
}

bool BiffFormulaParserImpl::importEndSheetToken3( BiffInputStream& rStrm )
{
    rStrm.ignore( 4 );
    mnCurrRefId = 0;
    return true;
}

bool BiffFormulaParserImpl::importNlrToken( BiffInputStream& rStrm )
{
    bool bOk = true;
    sal_uInt8 nNlrType;
    rStrm >> nNlrType;
    switch( nNlrType )
    {
        case BIFF_TOK_NLR_ERR:      bOk = importNlrErrToken( rStrm, 4 );        break;
        case BIFF_TOK_NLR_ROWR:     bOk = importNlrAddrToken( rStrm, true );    break;
        case BIFF_TOK_NLR_COLR:     bOk = importNlrAddrToken( rStrm, false );   break;
        case BIFF_TOK_NLR_ROWV:     bOk = importNlrAddrToken( rStrm, true );    break;
        case BIFF_TOK_NLR_COLV:     bOk = importNlrAddrToken( rStrm, false );   break;
        case BIFF_TOK_NLR_RANGE:    bOk = importNlrRangeToken( rStrm );         break;
        case BIFF_TOK_NLR_SRANGE:   bOk = importNlrSRangeToken( rStrm );        break;
        case BIFF_TOK_NLR_SROWR:    bOk = importNlrSAddrToken( rStrm, true );   break;
        case BIFF_TOK_NLR_SCOLR:    bOk = importNlrSAddrToken( rStrm, false );  break;
        case BIFF_TOK_NLR_SROWV:    bOk = importNlrSAddrToken( rStrm, true );   break;
        case BIFF_TOK_NLR_SCOLV:    bOk = importNlrSAddrToken( rStrm, false );  break;
        case BIFF_TOK_NLR_RANGEERR: bOk = importNlrErrToken( rStrm, 13 );       break;
        case BIFF_TOK_NLR_SXNAME:   bOk = importNlrErrToken( rStrm, 4 );        break;
        default:                    bOk = false;
    }
    return bOk;
}

bool BiffFormulaParserImpl::importArrayToken( BiffInputStream& rStrm )
{
    rStrm.ignore( mnArraySize );

    // start token array with opening brace and leading spaces
    pushOperand( mrFuncProv.OPCODE_ARRAY_OPEN );
    size_t nOpSize = popOperandSize();
    size_t nOldArraySize = maTokenIndexes.size();
    bool bBiff8 = getBiff() == BIFF8;

    // read array size
    swapStreamPosition( rStrm );
    sal_uInt16 nCols = rStrm.readuInt8();
    sal_uInt16 nRows = rStrm.readuInt16();
    if( bBiff8 ) { ++nCols; ++nRows; } else if( nCols == 0 ) nCols = 256;
    OSL_ENSURE( (nCols > 0) && (nRows > 0), "BiffFormulaParserImpl::importArrayToken - empty array" );

    // read array values and build token array
    for( sal_uInt16 nRow = 0; rStrm.isValid() && (nRow < nRows); ++nRow )
    {
        if( nRow > 0 )
            appendRawToken( mrFuncProv.OPCODE_ARRAY_ROWSEP );
        for( sal_uInt16 nCol = 0; rStrm.isValid() && (nCol < nCols); ++nCol )
        {
            if( nCol > 0 )
                appendRawToken( mrFuncProv.OPCODE_ARRAY_COLSEP );
            switch( rStrm.readuInt8() )
            {
                case BIFF_DATATYPE_EMPTY:
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= OUString();
                    rStrm.ignore( 8 );
                break;
                case BIFF_DATATYPE_DOUBLE:
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= rStrm.readDouble();
                break;
                case BIFF_DATATYPE_STRING:
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= bBiff8 ?
                        rStrm.readUniString() :
                        rStrm.readByteString( false, getTextEncoding() );
                break;
                case BIFF_DATATYPE_BOOL:
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= static_cast< double >( (rStrm.readuInt8() == 0) ? 0.0 : 1.0 );
                    rStrm.ignore( 7 );
                break;
                case BIFF_DATATYPE_ERROR:
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= BiffHelper::calcDoubleFromError( rStrm.readuInt8() );
                    rStrm.ignore( 7 );
                break;
                default:
                    OSL_ENSURE( false, "BiffFormulaParserImpl::importArrayToken - unknown data type" );
                    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= BiffHelper::calcDoubleFromError( BIFF_ERR_NA );
            }
        }
    }
    swapStreamPosition( rStrm );

    // close token array and set resulting operand size
    appendRawToken( mrFuncProv.OPCODE_ARRAY_CLOSE );
    pushOperandSize( nOpSize + maTokenIndexes.size() - nOldArraySize );
    return true;
}

bool BiffFormulaParserImpl::importNameToken( BiffInputStream& rStrm )
{
    sal_uInt16 nNameId = readNameId( rStrm );
    return (mnCurrRefId > 0) ? pushBiffExtName( mnCurrRefId, nNameId ) : pushBiffName( nNameId );
}

bool BiffFormulaParserImpl::importRefToken2( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    BiffSingleRef2d aRef;
    aRef.readBiff2Data( rStrm, bRelativeAsOffset );
    return pushBiffSingleRef2d( aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importRefToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    BiffSingleRef2d aRef;
    aRef.readBiff8Data( rStrm, bRelativeAsOffset );
    return pushBiffSingleRef2d( aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importAreaToken2( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    BiffComplexRef2d aRef;
    aRef.readBiff2Data( rStrm, bRelativeAsOffset );
    return pushBiffComplexRef2d( aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importAreaToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    BiffComplexRef2d aRef;
    aRef.readBiff8Data( rStrm, bRelativeAsOffset );
    return pushBiffComplexRef2d( aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importRef3dToken5( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    LinkSheetRange aSheetRange = readSheetRange5( rStrm );
    BiffSingleRef2d aRef;
    aRef.readBiff2Data( rStrm, bRelativeAsOffset );
    return pushBiffSingleRef3d( aSheetRange, aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importRef3dToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    LinkSheetRange aSheetRange = readSheetRange8( rStrm );
    BiffSingleRef2d aRef;
    aRef.readBiff8Data( rStrm, bRelativeAsOffset );
    return pushBiffSingleRef3d( aSheetRange, aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importArea3dToken5( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    LinkSheetRange aSheetRange = readSheetRange5( rStrm );
    BiffComplexRef2d aRef;
    aRef.readBiff2Data( rStrm, bRelativeAsOffset );
    return pushBiffComplexRef3d( aSheetRange, aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importArea3dToken8( BiffInputStream& rStrm, bool bDeleted, bool bRelativeAsOffset )
{
    LinkSheetRange aSheetRange = readSheetRange8( rStrm );
    BiffComplexRef2d aRef;
    aRef.readBiff8Data( rStrm, bRelativeAsOffset );
    return pushBiffComplexRef3d( aSheetRange, aRef, bDeleted, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::importMemAreaToken( BiffInputStream& rStrm, bool bAddData )
{
    rStrm.ignore( mnMemAreaSize );
    if( bAddData )
        ignoreMemAreaAddData( rStrm );
    return true;
}

bool BiffFormulaParserImpl::importMemFuncToken( BiffInputStream& rStrm )
{
    rStrm.ignore( mnMemFuncSize );
    return true;
}

bool BiffFormulaParserImpl::importNameXToken( BiffInputStream& rStrm )
{
    sal_Int32 nRefId = readRefId( rStrm );
    sal_uInt16 nNameId = readNameId( rStrm );
    return pushBiffExtName( nRefId, nNameId );
}

bool BiffFormulaParserImpl::importFuncToken2( BiffInputStream& rStrm )
{
    sal_uInt8 nFuncId;
    rStrm >> nFuncId;
    return pushBiffFunction( nFuncId );
}

bool BiffFormulaParserImpl::importFuncToken4( BiffInputStream& rStrm )
{
    sal_uInt16 nFuncId;
    rStrm >> nFuncId;
    return pushBiffFunction( nFuncId );
}

bool BiffFormulaParserImpl::importFuncVarToken2( BiffInputStream& rStrm )
{
    sal_uInt8 nParamCount, nFuncId;
    rStrm >> nParamCount >> nFuncId;
    return pushBiffFunction( nFuncId, nParamCount );
}

bool BiffFormulaParserImpl::importFuncVarToken4( BiffInputStream& rStrm )
{
    sal_uInt8 nParamCount;
    sal_uInt16 nFuncId;
    rStrm >> nParamCount >> nFuncId;
    return pushBiffFunction( nFuncId, nParamCount & BIFF_TOK_FUNCVAR_COUNTMASK );
}

bool BiffFormulaParserImpl::importFuncCEToken( BiffInputStream& rStrm )
{
    rStrm.ignore( 2 );
    return pushBiffError( BIFF_ERR_NA );
}

bool BiffFormulaParserImpl::importExpToken2( BiffInputStream& rStrm )
{
    BiffAddress aBaseAddr;
    aBaseAddr.read( rStrm, false );
    return processExpToken( aBaseAddr );
}

bool BiffFormulaParserImpl::importExpToken3( BiffInputStream& rStrm )
{
    BiffAddress aBaseAddr;
    aBaseAddr.read( rStrm );
    return processExpToken( aBaseAddr );
}

bool BiffFormulaParserImpl::importTblToken2( BiffInputStream& rStrm )
{
    BiffAddress aBaseAddr;
    aBaseAddr.read( rStrm, false );
    return processTblToken( aBaseAddr );
}

bool BiffFormulaParserImpl::importTblToken3( BiffInputStream& rStrm )
{
    BiffAddress aBaseAddr;
    aBaseAddr.read( rStrm );
    return processTblToken( aBaseAddr );
}

bool BiffFormulaParserImpl::importNlrAddrToken( BiffInputStream& rStrm, bool bRow )
{
    BiffNlr aBiffNlr;
    aBiffNlr.readBiff8Data( rStrm );
    return pushBiffNlrAddr( aBiffNlr, bRow );
}

bool BiffFormulaParserImpl::importNlrRangeToken( BiffInputStream& rStrm )
{
    BiffNlr aBiffNlr;
    aBiffNlr.readBiff8Data( rStrm );
    rStrm.ignore( 1 );
    BiffRange aBiffRange;
    rStrm >> aBiffRange;
    return pushBiffNlrRange( aBiffNlr, aBiffRange );
}

bool BiffFormulaParserImpl::importNlrSAddrToken( BiffInputStream& rStrm, bool bRow )
{
    rStrm.ignore( 4 );
    BiffNlr aBiffNlr;
    return readNlrSAddrAddData( aBiffNlr, rStrm, bRow ) ? pushBiffNlrSAddr( aBiffNlr, bRow ) : pushBiffError( BIFF_ERR_REF );
}

bool BiffFormulaParserImpl::importNlrSRangeToken( BiffInputStream& rStrm )
{
    rStrm.ignore( 5 );
    BiffRange aBiffRange;
    rStrm >> aBiffRange;
    BiffNlr aBiffNlr;
    bool bRow;
    return readNlrSRangeAddData( aBiffNlr, bRow, rStrm ) ? pushBiffNlrSRange( aBiffNlr, aBiffRange, bRow ) : pushBiffError( BIFF_ERR_REF );
}

bool BiffFormulaParserImpl::importNlrErrToken( BiffInputStream& rStrm, sal_uInt16 nIgnore )
{
    rStrm.ignore( nIgnore );
    return pushBiffError( BIFF_ERR_NAME );
}

sal_Int32 BiffFormulaParserImpl::readRefId( BiffInputStream& rStrm )
{
    sal_Int16 nRefId;
    rStrm >> nRefId;
    rStrm.ignore( mnRefIdSize );
    return nRefId;
}

sal_uInt16 BiffFormulaParserImpl::readNameId( BiffInputStream& rStrm )
{
    sal_uInt16 nNameId;
    rStrm >> nNameId;
    rStrm.ignore( mnNameSize );
    return nNameId;
}

LinkSheetRange BiffFormulaParserImpl::readSheetRange5( BiffInputStream& rStrm )
{
    sal_Int32 nRefId = readRefId( rStrm );
    sal_Int16 nTab1 = 0, nTab2 = 0;
    // internal references (negative ref id) - sheets are specified seprately
    if( nRefId < 0 )
        rStrm >> nTab1 >> nTab2;
    else
        rStrm.ignore( 4 );
    return getExternalLinks().getSheetRange( nRefId, nTab1, nTab2 );
}

LinkSheetRange BiffFormulaParserImpl::readSheetRange8( BiffInputStream& rStrm )
{
    return getExternalLinks().getSheetRange( readRefId( rStrm ) );
}

void BiffFormulaParserImpl::swapStreamPosition( BiffInputStream& rStrm )
{
    sal_uInt32 nRecPos = rStrm.getRecPos();
    rStrm.seek( mnAddDataPos );
    mnAddDataPos = nRecPos;
}

void BiffFormulaParserImpl::ignoreMemAreaAddData( BiffInputStream& rStrm )
{
    swapStreamPosition( rStrm );
    sal_uInt32 nCount = rStrm.readuInt16();
    rStrm.ignore( ((getBiff() == BIFF8) ? 8 : 6) * nCount );
    swapStreamPosition( rStrm );
}

bool BiffFormulaParserImpl::readNlrSAddrAddData( BiffNlr& orBiffNlr, BiffInputStream& rStrm, bool bRow )
{
    bool bIsRow;
    return readNlrSRangeAddData( orBiffNlr, bIsRow, rStrm ) && (bIsRow == bRow);
}

bool BiffFormulaParserImpl::readNlrSRangeAddData( BiffNlr& orBiffNlr, bool& orbIsRow, BiffInputStream& rStrm )
{
    swapStreamPosition( rStrm );
    // read number of cell addresses and relative flag
    sal_uInt32 nCount;
    rStrm >> nCount;
    bool bRel = getFlag( nCount, BIFF_TOK_NLR_ADDREL );
    nCount &= BIFF_TOK_NLR_ADDMASK;
    sal_uInt32 nEndPos = rStrm.getRecPos() + 4 * nCount;
    // read list of cell addresses
    bool bValid = false;
    if( nCount >= 2 )
    {
        // detect column/row orientation
        BiffAddress aBiffAddr1, aBiffAddr2;
        rStrm >> aBiffAddr1 >> aBiffAddr2;
        orbIsRow = aBiffAddr1.mnRow == aBiffAddr2.mnRow;
        bValid = lclIsValidNlrStack( aBiffAddr1, aBiffAddr2, orbIsRow );
        // read and verify additional cell positions
        for( sal_uInt32 nIndex = 2; bValid && (nIndex < nCount); ++nIndex )
        {
            aBiffAddr1 = aBiffAddr2;
            rStrm >> aBiffAddr2;
            bValid = rStrm.isValid() && lclIsValidNlrStack( aBiffAddr1, aBiffAddr2, orbIsRow );
        }
        // check that last imported position (aBiffAddr2) is not at the end of the sheet
        bValid = bValid && (orbIsRow ? (aBiffAddr2.mnCol < mnMaxCol) : (aBiffAddr2.mnRow < mnMaxRow));
        // fill the NLR struct with the last imported position
        if( bValid )
        {
            orBiffNlr.mnCol = aBiffAddr2.mnCol;
            orBiffNlr.mnRow = aBiffAddr2.mnRow;
            orBiffNlr.mbRel = bRel;
        }
    }
    // seek to end of additional data for this token
    rStrm.seek( nEndPos );
    swapStreamPosition( rStrm );

    return bValid;
}

// push API operand or operator -----------------------------------------------

bool BiffFormulaParserImpl::pushOperandToken( sal_Int32 nOpCode, sal_Int32 nSpaces )
{
    size_t nSpacesSize = appendSpacesToken( nSpaces );
    appendRawToken( nOpCode );
    pushOperandSize( nSpacesSize + 1 );
    return true;
}

template< typename Type >
bool BiffFormulaParserImpl::pushValueOperandToken( const Type& rValue, sal_Int32 nOpCode, sal_Int32 nSpaces )
{
    size_t nSpacesSize = appendSpacesToken( nSpaces );
    appendRawToken( nOpCode ) <<= rValue;
    pushOperandSize( nSpacesSize + 1 );
    return true;
}

bool BiffFormulaParserImpl::pushParenthesesOperandToken( sal_Int32 nOpeningSpaces, sal_Int32 nClosingSpaces )
{
    size_t nSpacesSize = appendSpacesToken( nOpeningSpaces );
    appendRawToken( mrFuncProv.OPCODE_OPEN );
    nSpacesSize += appendSpacesToken( nClosingSpaces );
    appendRawToken( mrFuncProv.OPCODE_CLOSE );
    pushOperandSize( nSpacesSize + 2 );
    return true;
}

bool BiffFormulaParserImpl::pushUnaryPreOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces )
{
    bool bOk = maOperandSizeStack.size() >= 1;
    if( bOk )
    {
        size_t nOpSize = popOperandSize();
        size_t nSpacesSize = insertSpacesToken( nSpaces, nOpSize );
        insertRawToken( nOpCode, nOpSize );
        pushOperandSize( nOpSize + nSpacesSize + 1 );
    }
    return bOk;
}

bool BiffFormulaParserImpl::pushUnaryPostOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces )
{
    bool bOk = maOperandSizeStack.size() >= 1;
    if( bOk )
    {
        size_t nOpSize = popOperandSize();
        size_t nSpacesSize = appendSpacesToken( nSpaces );
        appendRawToken( nOpCode );
        pushOperandSize( nOpSize + nSpacesSize + 1 );
    }
    return bOk;
}

bool BiffFormulaParserImpl::pushBinaryOperatorToken( sal_Int32 nOpCode, sal_Int32 nSpaces )
{
    bool bOk = maOperandSizeStack.size() >= 2;
    if( bOk )
    {
        size_t nOp2Size = popOperandSize();
        size_t nOp1Size = popOperandSize();
        size_t nSpacesSize = insertSpacesToken( nSpaces, nOp2Size );
        insertRawToken( nOpCode, nOp2Size );
        pushOperandSize( nOp1Size + nSpacesSize + 1 + nOp2Size );
    }
    return bOk;
}

bool BiffFormulaParserImpl::pushParenthesesOperatorToken( sal_Int32 nOpeningSpaces, sal_Int32 nClosingSpaces )
{
    bool bOk = !maOperandSizeStack.empty();
    if( bOk )
    {
        size_t nOpSize = popOperandSize();
        size_t nSpacesSize = insertSpacesToken( nOpeningSpaces, nOpSize );
        insertRawToken( mrFuncProv.OPCODE_OPEN, nOpSize );
        nSpacesSize += appendSpacesToken( nClosingSpaces );
        appendRawToken( mrFuncProv.OPCODE_CLOSE );
        pushOperandSize( nOpSize + nSpacesSize + 2 );
    }
    return bOk;
}

bool BiffFormulaParserImpl::pushOperand( sal_Int32 nOpCode )
{
    return pushOperandToken( nOpCode, mnLeadingSpaces ) && resetSpaces();
}

template< typename Type >
bool BiffFormulaParserImpl::pushValueOperand( const Type& rValue, sal_Int32 nOpCode )
{
    return pushValueOperandToken( rValue, nOpCode, mnLeadingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushParenthesesOperand()
{
    return pushParenthesesOperandToken( mnOpeningSpaces, mnClosingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushDelAddressOperand()
{
    SingleReference aRef;
    initializeRef2d( aRef );
    convertColRowDel( aRef );
    return pushValueOperand( aRef );
}

bool BiffFormulaParserImpl::pushUnaryPreOperator( sal_Int32 nOpCode )
{
    return pushUnaryPreOperatorToken( nOpCode, mnLeadingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushUnaryPostOperator( sal_Int32 nOpCode )
{
    return pushUnaryPostOperatorToken( nOpCode, mnLeadingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushBinaryOperator( sal_Int32 nOpCode )
{
    return pushBinaryOperatorToken( nOpCode, mnLeadingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushRangeOperator()
{
    // #i48496# try to convert the term SingleRef:SingleRef to a ComplexReference
    bool bOk = maOperandSizeStack.size() >= 2;
    if( bOk )
    {
        bool bConvertedToRange = false;
        if( (getOperandSize( 2, 0 ) == 1) && (getOperandSize( 2, 1 ) == 1) )
        {
            ApiToken& rOp1 = getOperandToken( 2, 0 );
            ApiToken& rOp2 = getOperandToken( 2, 1 );
            if( (rOp1.OpCode == mrFuncProv.OPCODE_PUSH) && (rOp2.OpCode == rOp1.OpCode) )
            {
                SingleReference aRef1, aRef2;
                if( (rOp1.Data >>= aRef1) && (rOp2.Data >>= aRef2) )
                {
                    removeLastOperands( 2 );
                    bOk = pushValueOperand( ComplexReference( aRef1, aRef2 ) );
                    bConvertedToRange = true;
                }
            }
        }
        if( !bConvertedToRange )
            bOk = pushBinaryOperator( mrFuncProv.OPCODE_RANGE );
    }
    return bOk;
}

bool BiffFormulaParserImpl::pushParenthesesOperator()
{
    return pushParenthesesOperatorToken( mnOpeningSpaces, mnClosingSpaces ) && resetSpaces();
}

bool BiffFormulaParserImpl::pushFunctionOperator( sal_Int32 nOpCode, size_t nParamCount )
{
    bool bOk = nParamCount <= maOperandSizeStack.size();

    // convert all parameters on stack to a single operand separated with OPCODE_SEP
    for( size_t nParam = 1; bOk && (nParam < nParamCount); ++nParam )
        bOk = pushBinaryOperatorToken( mrFuncProv.OPCODE_SEP, 0 );

    // add function parentheses and function name
    return bOk &&
        ((nParamCount > 0) ? pushParenthesesOperatorToken( 0, mnClosingSpaces ) : pushParenthesesOperandToken( 0, mnClosingSpaces )) &&
        pushUnaryPreOperatorToken( nOpCode, mnLeadingSpaces ) && resetSpaces();
}

size_t BiffFormulaParserImpl::getOperandSize( size_t nOpCountFromEnd, size_t nOpIndex )
{
    OSL_ENSURE( (nOpIndex < nOpCountFromEnd) && (nOpCountFromEnd <= maOperandSizeStack.size()),
        "BiffFormulaParserImpl::getOperandSize - invalid parameters" );
    return maOperandSizeStack[ maOperandSizeStack.size() - nOpCountFromEnd + nOpIndex ];
}

void BiffFormulaParserImpl::pushOperandSize( size_t nSize )
{
    maOperandSizeStack.push_back( nSize );
}

size_t BiffFormulaParserImpl::popOperandSize()
{
    OSL_ENSURE( !maOperandSizeStack.empty(), "BiffFormulaParserImpl::popOperandSize - invalid call" );
    size_t nOpSize = maOperandSizeStack.back();
    maOperandSizeStack.pop_back();
    return nOpSize;
}

ApiToken& BiffFormulaParserImpl::getOperandToken( size_t nOpCountFromEnd, size_t nOpIndex )
{
    OSL_ENSURE( (nOpIndex < nOpCountFromEnd) && (nOpCountFromEnd <= maOperandSizeStack.size()),
        "BiffFormulaParserImpl::getOperandToken - invalid parameters" );
    SizeTypeVec::const_iterator aIndexIt = maTokenIndexes.end();
    for( SizeTypeVec::const_iterator aEnd = maOperandSizeStack.end(), aIt = aEnd - nOpCountFromEnd + nOpIndex; aIt != aEnd; ++aIt )
        aIndexIt -= *aIt;
    return maTokenStorage[ *aIndexIt ];
}

void BiffFormulaParserImpl::removeOperand( size_t nOpCountFromEnd, size_t nOpIndex )
{
    OSL_ENSURE( (nOpIndex < nOpCountFromEnd) && (nOpCountFromEnd <= maOperandSizeStack.size()),
        "BiffFormulaParserImpl::removeOperand - invalid parameters" );
    // remove indexes into token storage, but do not touch storage itself
    SizeTypeVec::iterator aSizeEnd = maOperandSizeStack.end();
    SizeTypeVec::iterator aSizeIt = aSizeEnd - nOpCountFromEnd + nOpIndex;
    size_t nRemainingSize = 0;
    for( SizeTypeVec::iterator aIt = aSizeIt + 1; aIt != aSizeEnd; ++aIt )
        nRemainingSize += *aIt;
    maTokenIndexes.erase( maTokenIndexes.end() - nRemainingSize - *aSizeIt, maTokenIndexes.end() - nRemainingSize );
    maOperandSizeStack.erase( aSizeIt );
}

void BiffFormulaParserImpl::removeLastOperands( size_t nOpCountFromEnd )
{
    for( size_t nOpIndex = 0; nOpIndex < nOpCountFromEnd; ++nOpIndex )
        removeOperand( 1, 0 );
}

Any& BiffFormulaParserImpl::appendRawToken( sal_Int32 nOpCode )
{
    size_t nTokenIndex = maTokenStorage.size();
    maTokenStorage.resize( nTokenIndex + 1 );
    maTokenStorage.back().OpCode = nOpCode;
    maTokenIndexes.push_back( nTokenIndex );
    return maTokenStorage.back().Data;
}

Any& BiffFormulaParserImpl::insertRawToken( sal_Int32 nOpCode, size_t nIndexFromEnd )
{
    size_t nTokenIndex = maTokenStorage.size();
    maTokenStorage.resize( nTokenIndex + 1 );
    maTokenStorage.back().OpCode = nOpCode;
    maTokenIndexes.insert( maTokenIndexes.end() - nIndexFromEnd, nTokenIndex );
    return maTokenStorage.back().Data;
}

size_t BiffFormulaParserImpl::appendSpacesToken( sal_Int32 nSpaces )
{
    if( nSpaces > 0 )
    {
        appendRawToken( mrFuncProv.OPCODE_SPACES ) <<= nSpaces;
        return 1;
    }
    return 0;
}

size_t BiffFormulaParserImpl::insertSpacesToken( sal_Int32 nSpaces, size_t nIndexFromEnd )
{
    if( nSpaces > 0 )
    {
        insertRawToken( mrFuncProv.OPCODE_SPACES, nIndexFromEnd ) <<= nSpaces;
        return 1;
    }
    return 0;
}

bool BiffFormulaParserImpl::resetSpaces()
{
    mnLeadingSpaces = mnOpeningSpaces = mnClosingSpaces = 0;
    return true;
}

// convert BIFF token and push API operand or operator ------------------------

bool BiffFormulaParserImpl::pushBiffBool( sal_uInt8 nValue )
{
    return pushBiffFunction( (nValue == BIFF_TOK_BOOL_FALSE) ? BIFF_FUNC_FALSE : BIFF_FUNC_TRUE, 0 );
}

bool BiffFormulaParserImpl::pushBiffError( sal_uInt8 nErrorCode )
{
    // HACK: enclose all error codes into an 1x1 matrix
    // start token array with opening brace and leading spaces
    pushOperand( mrFuncProv.OPCODE_ARRAY_OPEN );
    size_t nOpSize = popOperandSize();
    size_t nOldArraySize = maTokenIndexes.size();
    // push a double containing the Calc error code
    appendRawToken( mrFuncProv.OPCODE_PUSH ) <<= BiffHelper::calcDoubleFromError( nErrorCode );
    // close token array and set resulting operand size
    appendRawToken( mrFuncProv.OPCODE_ARRAY_CLOSE );
    pushOperandSize( nOpSize + maTokenIndexes.size() - nOldArraySize );
    return true;
}

bool BiffFormulaParserImpl::pushBiffSingleRef2d( const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset )
{
    if( mnCurrRefId > 0 )
        return pushBiffSingleRef3d( getExternalLinks().getSheetRange( mnCurrRefId ), rBiffRef, bDeleted, bRelativeAsOffset );
    SingleReference aApiRef;
    convertSingleRef2d( aApiRef, rBiffRef, bDeleted, bRelativeAsOffset );
    return pushValueOperand( aApiRef );
}

bool BiffFormulaParserImpl::pushBiffComplexRef2d( const BiffComplexRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset )
{
    if( mnCurrRefId > 0 )
        return pushBiffComplexRef3d( getExternalLinks().getSheetRange( mnCurrRefId ), rBiffRef, bDeleted, bRelativeAsOffset );
    ComplexReference aApiRef;
    convertSingleRef2d( aApiRef.Reference1, rBiffRef.maRef1, bDeleted, bRelativeAsOffset );
    convertSingleRef2d( aApiRef.Reference2, rBiffRef.maRef2, bDeleted, bRelativeAsOffset );
    return pushValueOperand( aApiRef );
}

bool BiffFormulaParserImpl::pushBiffSingleRef3d( const LinkSheetRange& rSheetRange, const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset )
{
    if( rSheetRange.is3dRange() )
    {
        // single-cell-range over several sheets, need to create a ComplexReference
        ComplexReference aApiRef;
        convertSingleRef3d( aApiRef.Reference1, rBiffRef, rSheetRange.mnFirst, bDeleted, bRelativeAsOffset );
        convertSingleRef3d( aApiRef.Reference2, rBiffRef, rSheetRange.mnLast, bDeleted, bRelativeAsOffset );
        return pushValueOperand( aApiRef );
    }
    sal_Int32 nSheet = rSheetRange.isDeleted() ? -1 : rSheetRange.mnFirst;
    SingleReference aApiRef;
    convertSingleRef3d( aApiRef, rBiffRef, nSheet, bDeleted, bRelativeAsOffset );
    return pushValueOperand( aApiRef );
}

bool BiffFormulaParserImpl::pushBiffComplexRef3d( const LinkSheetRange& rSheetRange, const BiffComplexRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset )
{
    sal_Int32 nFirstSheet = rSheetRange.isDeleted() ? -1 : rSheetRange.mnFirst;
    sal_Int32 nLastSheet = rSheetRange.isDeleted() ? -1 : rSheetRange.mnLast;
    ComplexReference aApiRef;
    convertSingleRef3d( aApiRef.Reference1, rBiffRef.maRef1, nFirstSheet, bDeleted, bRelativeAsOffset );
    convertSingleRef3d( aApiRef.Reference2, rBiffRef.maRef2, nLastSheet, bDeleted, bRelativeAsOffset );
    // remove sheet name from second part of reference
    setFlag( aApiRef.Reference2.Flags, SHEET_3D, rSheetRange.is3dRange() );
    return pushValueOperand( aApiRef );
}

bool BiffFormulaParserImpl::pushBiffNlrAddr( const BiffNlr& rBiffNlr, bool bRow )
{
    BiffSingleRef2d aBiffRef;
    aBiffRef.mnCol = rBiffNlr.mnCol;
    aBiffRef.mnRow = rBiffNlr.mnRow;
    aBiffRef.mbColRel = !bRow;
    aBiffRef.mbRowRel = bRow;
    SingleReference aApiRef;
    convertSingleRef2d( aApiRef, aBiffRef, false, false );
    return pushValueOperand( aApiRef, mrFuncProv.OPCODE_NLR );
}

bool BiffFormulaParserImpl::pushBiffNlrRange( const BiffNlr& rBiffNlr, const BiffRange& rBiffRange )
{
    bool bRow = rBiffNlr.mnRow == rBiffRange.maFirst.mnRow;
    return lclIsValidNlrRange( rBiffNlr, rBiffRange, bRow ) ?
        pushBiffNlrAddr( rBiffNlr, bRow ) : pushBiffError( BIFF_ERR_REF );
}

bool BiffFormulaParserImpl::pushBiffNlrSAddr( const BiffNlr& rBiffNlr, bool bRow )
{
    BiffRange aBiffRange;
    aBiffRange.maFirst.mnCol = static_cast< sal_uInt16 >( rBiffNlr.mnCol + (bRow ? 1 : 0) );
    aBiffRange.maFirst.mnRow = static_cast< sal_uInt16 >( rBiffNlr.mnRow + (bRow ? 0 : 1) );
    aBiffRange.maLast.mnCol = bRow ? mnMaxCol : rBiffNlr.mnCol;
    aBiffRange.maLast.mnRow = bRow ? rBiffNlr.mnRow : mnMaxRow;
    return pushBiffNlrSRange( rBiffNlr, aBiffRange, bRow );
}

bool BiffFormulaParserImpl::pushBiffNlrSRange( const BiffNlr& rBiffNlr, const BiffRange& rBiffRange, bool bRow )
{
    if( lclIsValidNlrRange( rBiffNlr, rBiffRange, bRow ) )
    {
        BiffComplexRef2d aBiffRef;
        aBiffRef.maRef1.mnCol = rBiffRange.maFirst.mnCol;
        aBiffRef.maRef1.mnRow = rBiffRange.maFirst.mnRow;
        aBiffRef.maRef2.mnCol = rBiffRange.maLast.mnCol;
        aBiffRef.maRef2.mnRow = rBiffRange.maLast.mnRow;
        aBiffRef.maRef1.mbColRel = aBiffRef.maRef2.mbColRel = !bRow && rBiffNlr.mbRel;
        aBiffRef.maRef1.mbRowRel = aBiffRef.maRef2.mbRowRel = bRow && rBiffNlr.mbRel;
        ComplexReference aApiRef;
        convertSingleRef2d( aApiRef.Reference1, aBiffRef.maRef1, false, false );
        convertSingleRef2d( aApiRef.Reference2, aBiffRef.maRef2, false, false );
        return pushValueOperand( aApiRef );
    }
    return pushBiffError( BIFF_ERR_REF );
}

bool BiffFormulaParserImpl::pushBiffName( sal_uInt16 nNameId )
{
    if( const DefinedName* pDefName = getDefinedNames().getByBiffIndex( nNameId ).get() )
    {
        if( pDefName->isMacroFunc( false ) )
            return pushValueOperand( pDefName->getOoxName(), mrFuncProv.OPCODE_MACRO );
        if( pDefName->getTokenIndex() >= 0 )
            return pushValueOperand( pDefName->getTokenIndex(), mrFuncProv.OPCODE_NAME );
    }
    return pushBiffError( BIFF_ERR_NAME );
}

bool BiffFormulaParserImpl::pushBiffExtName( sal_Int32 nRefId, sal_uInt16 nNameId )
{
    switch( getExternalLinks().getLinkType( nRefId ) )
    {
        case LINKTYPE_INTERNAL: return pushBiffName( nNameId );
        case LINKTYPE_ANALYSIS: return pushBiffAnalysisName( nRefId, nNameId );
        default:;
    }
    return pushBiffError( BIFF_ERR_NAME );
}

bool BiffFormulaParserImpl::pushBiffAnalysisName( sal_Int32 nRefId, sal_uInt16 nNameId )
{
    if( const ExternalName* pExtName = getExternalLinks().getExternalName( nRefId, nNameId ).get() )
        if( const FunctionInfo* pFuncInfo = mrFuncProv.getFuncInfoFromOoxFuncName( pExtName->getName() ) )
            if( pFuncInfo->maExternCallName.getLength() > 0 )
                return pushValueOperand( pFuncInfo->maExternCallName, mrFuncProv.OPCODE_MACRO );
    return pushBiffError( BIFF_ERR_NAME );
}

bool BiffFormulaParserImpl::pushBiffFunction( const FunctionInfo& rFuncInfo, sal_uInt8 nParamCount )
{
    /*  #i70925# if there are not enough tokens available on token stack, do
        not exit with error, but reduce parameter count. */
    size_t nRealParamCount = ::std::min< size_t >( maOperandSizeStack.size(), nParamCount );
    return pushFunctionOperator( rFuncInfo.mnApiOpCode, nRealParamCount );
}

bool BiffFormulaParserImpl::pushBiffFunction( sal_uInt16 nFuncId )
{
    if( const FunctionInfo* pFuncInfo = mrFuncProv.getFuncInfoFromBiffFuncId( nFuncId ) )
        if( pFuncInfo->mnMinParamCount == pFuncInfo->mnMaxParamCount )
            return pushBiffFunction( *pFuncInfo, pFuncInfo->mnMinParamCount );
    return pushFunctionOperator( mrFuncProv.OPCODE_NONAME, 0 );
}

bool BiffFormulaParserImpl::pushBiffFunction( sal_uInt16 nFuncId, sal_uInt8 nParamCount )
{
    if( const FunctionInfo* pFuncInfo = mrFuncProv.getFuncInfoFromBiffFuncId( nFuncId ) )
        return pushBiffFunction( *pFuncInfo, nParamCount );
    return pushFunctionOperator( mrFuncProv.OPCODE_NONAME, nParamCount );
}

void BiffFormulaParserImpl::initializeRef2d( SingleReference& orApiRef ) const
{
    orApiRef.Flags = SHEET_RELATIVE;
    // #i10184# absolute sheet index needed for relative references in shared formulas
    orApiRef.Sheet = getFormulaContext().getBaseAddress().Sheet;
    orApiRef.RelativeSheet = 0;
}

void BiffFormulaParserImpl::initializeRef3d( SingleReference& orApiRef, sal_Int32 nSheet ) const
{
    orApiRef.Flags = SHEET_3D;
    if( nSheet < 0 )
    {
        orApiRef.Sheet = 0;
        orApiRef.Flags |= SHEET_DELETED;
    }
    else
        orApiRef.Sheet = nSheet;
}

void BiffFormulaParserImpl::convertColRow( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, bool bRelativeAsOffset ) const
{
    // column/row indexes and flags
    setFlag( orApiRef.Flags, COLUMN_RELATIVE, rBiffRef.mbColRel );
    setFlag( orApiRef.Flags, ROW_RELATIVE, rBiffRef.mbRowRel );
    (rBiffRef.mbColRel ? orApiRef.RelativeColumn : orApiRef.Column) = rBiffRef.mnCol;
    (rBiffRef.mbRowRel ? orApiRef.RelativeRow : orApiRef.Row) = rBiffRef.mnRow;
    // convert absolute BIFF indexes to relative offsets used in API
    if( !bRelativeAsOffset )
    {
        if( rBiffRef.mbColRel )
            orApiRef.RelativeColumn -= getFormulaContext().getBaseAddress().Column;
        if( rBiffRef.mbRowRel )
            orApiRef.RelativeRow -= getFormulaContext().getBaseAddress().Row;
    }
}

void BiffFormulaParserImpl::convertColRowDel( SingleReference& orApiRef ) const
{
    orApiRef.Column = 0;
    orApiRef.Row = 0;
    orApiRef.Flags |= COLUMN_DELETED | ROW_DELETED;
}

void BiffFormulaParserImpl::convertSingleRef2d( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, bool bDeleted, bool bRelativeAsOffset ) const
{
    initializeRef2d( orApiRef );
    if( bDeleted )
        convertColRowDel( orApiRef );
    else
        convertColRow( orApiRef, rBiffRef, bRelativeAsOffset );
}

void BiffFormulaParserImpl::convertSingleRef3d( SingleReference& orApiRef, const BiffSingleRef2d& rBiffRef, sal_Int32 nSheet, bool bDeleted, bool bRelativeAsOffset ) const
{
    initializeRef3d( orApiRef, nSheet );
    if( bDeleted )
        convertColRowDel( orApiRef );
    else
        convertColRow( orApiRef, rBiffRef, bRelativeAsOffset );
}

bool BiffFormulaParserImpl::processExpToken( const BiffAddress& rBaseAddr )
{
    CellAddress aAddr;
    if( getAddressConverter().convertToCellAddress( aAddr, rBaseAddr, 0, false ) )
        getFormulaContext().setSharedFormula( AddressConverter::convertToId( aAddr ) );
    // formula has been set, exit parser by returning false
    return false;
}

bool BiffFormulaParserImpl::processTblToken( const BiffAddress& rBaseAddr )
{
    CellAddress aAddr;
    if( getAddressConverter().convertToCellAddress( aAddr, rBaseAddr, 0, false ) )
        getFormulaContext().setTableOperation( AddressConverter::convertToId( aAddr ) );
    // formula has been set, exit parser by returning false
    return false;
}

// ============================================================================

FormulaParser::FormulaParser( const GlobalDataHelper& rGlobalData ) :
    FormulaProcessorBase( rGlobalData )
{
    switch( getFilterType() )
    {
        case FILTER_OOX:    mxImpl.reset( new OoxFormulaParserImpl( rGlobalData, maFuncProv ) );    break;
        case FILTER_BIFF:   mxImpl.reset( new BiffFormulaParserImpl( rGlobalData, maFuncProv ) );   break;
        case FILTER_UNKNOWN: break;
    }
}

FormulaParser::~FormulaParser()
{
}

void FormulaParser::importFormula( FormulaContext& rContext, const OUString& rFormulaString ) const
{
    mxImpl->importOoxFormula( rContext, rFormulaString );
}

void FormulaParser::importFormula( FormulaContext& rContext, BiffInputStream& rStrm, const sal_uInt16* pnFmlaSize ) const
{
    mxImpl->importBiffFormula( rContext, rStrm, pnFmlaSize );
}

void FormulaParser::convertErrorToFormula( FormulaContext& rContext, sal_uInt8 nErrorCode ) const
{
    ApiTokenSequence aTokens( 3 );
    // HACK: enclose all error codes into an 1x1 matrix
    aTokens[ 0 ].OpCode = maFuncProv.OPCODE_ARRAY_OPEN;
    aTokens[ 1 ].OpCode = maFuncProv.OPCODE_PUSH;
    aTokens[ 1 ].Data <<= BiffHelper::calcDoubleFromError( nErrorCode );
    aTokens[ 2 ].OpCode = maFuncProv.OPCODE_ARRAY_CLOSE;
    mxImpl->setFormula( rContext, aTokens );
}

void FormulaParser::convertNameToFormula( FormulaContext& rContext, sal_Int32 nTokenIndex ) const
{
    if( nTokenIndex >= 0 )
    {
        ApiTokenSequence aTokens( 1 );
        aTokens[ 0 ].OpCode = maFuncProv.OPCODE_NAME;
        aTokens[ 0 ].Data <<= nTokenIndex;
        mxImpl->setFormula( rContext, aTokens );
    }
    else
        convertErrorToFormula( rContext, BIFF_ERR_REF );
}

// ============================================================================

} // namespace xls
} // namespace oox

