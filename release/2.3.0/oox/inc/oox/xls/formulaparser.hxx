/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: formulaparser.hxx,v $
 *
 *  $Revision: 1.1.2.20 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/16 16:26:37 $
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

#ifndef OOX_XLS_FORMULAPARSER_HXX
#define OOX_XLS_FORMULAPARSER_HXX

#include <com/sun/star/table/CellAddress.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include "oox/xls/formulabase.hxx"

namespace com { namespace sun { namespace star {
    namespace sheet {
        class XFormulaTokens;
        class XMultiFormulaTokens;
        class XArrayFormulaTokens;
    }
} } }

namespace oox {
namespace xls {

// ============================================================================

class FormulaContext
{
public:
    void                setBaseAddress(
                            const ::com::sun::star::table::CellAddress& rBaseAddress,
                            bool bRelativeAsOffset = false );

    inline const ::com::sun::star::table::CellAddress& getBaseAddress() const { return maBaseAddress; }
    inline bool         isRelativeAsOffset() const { return mbRelativeAsOffset; }

    virtual void        setTokens( const ApiTokenSequence& rTokens ) = 0;
    virtual void        setSharedFormula( sal_Int32 nSharedId );
    virtual void        setTableOperation( sal_Int32 nTableOpId );

protected:
    explicit            FormulaContext();
    virtual             ~FormulaContext();

private:
    ::com::sun::star::table::CellAddress maBaseAddress;
    bool                mbRelativeAsOffset;
};

// ----------------------------------------------------------------------------

/** Stores the converted formula token sequence into the passed external Any. */
class AnyFormulaContext : public FormulaContext
{
public:
    explicit            AnyFormulaContext( ::com::sun::star::uno::Any& rAny );

    virtual void        setTokens( const ApiTokenSequence& rTokens );

private:
    ::com::sun::star::uno::Any& mrAny;
};

// ----------------------------------------------------------------------------

/** Uses the XFormulaTokens interface to set a token sequence. */
class SimpleFormulaContext : public FormulaContext
{
public:
    explicit            SimpleFormulaContext(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XFormulaTokens >& rxTokens );

    virtual void        setTokens( const ApiTokenSequence& rTokens );

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XFormulaTokens > mxTokens;
};

// ----------------------------------------------------------------------------

/** Uses the XMultiFormulaTokens interface to set a token sequence. */
class MultiFormulaContext : public FormulaContext
{
public:
    explicit            MultiFormulaContext(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XMultiFormulaTokens >& rxTokens,
                            sal_Int32 nIndex );

    virtual void        setTokens( const ApiTokenSequence& rTokens );

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XMultiFormulaTokens > mxTokens;
    sal_Int32           mnIndex;
};

// ----------------------------------------------------------------------------

/** Uses the XArrayFormulaTokens interface to set a token sequence. */
class ArrayFormulaContext : public FormulaContext
{
public:
    explicit            ArrayFormulaContext(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XArrayFormulaTokens >& rxTokens,
                            const ::com::sun::star::table::CellRangeAddress& rArrayRange );

    virtual void        setTokens( const ApiTokenSequence& rTokens );

private:
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XArrayFormulaTokens > mxTokens;
};

// ============================================================================

class FormulaParserImpl;

/** Import formula parser for OOX and BIFF filters.

    This class implements formula import for the OOX and BIFF filter. One
    instance is contained in the global filter data to prevent construction and
    destruction of internal buffers for every imported formula.
 */
class FormulaParser : public FormulaProcessorBase
{
public:
    explicit            FormulaParser( const GlobalDataHelper& rGlobalData );
    virtual             ~FormulaParser();

    /** Converts an XML formula string. */
    void                importFormula(
                            FormulaContext& rContext,
                            const ::rtl::OUString& rFormulaString ) const;

    /** Imports and converts a BIFF token array from the passed stream.
        @param pnFmlaSize  Size of the token array. If 0 is passed, reads
        it from stream (1 byte in BIFF2, 2 bytes otherwise) first. */
    void                importFormula(
                            FormulaContext& rContext,
                            BiffInputStream& rStrm,
                            const sal_uInt16* pnFmlaSize = 0 ) const;

    /** Converts the passed BIFF error code to a similar formula. */
    void                convertErrorToFormula(
                            FormulaContext& rContext,
                            sal_uInt8 nErrorCode ) const;

    /** Converts the passed token index of a defined name to a formula calling that name. */
    void                convertNameToFormula(
                            FormulaContext& rContext,
                            sal_Int32 nTokenIndex ) const;

private:
    ::std::auto_ptr< FormulaParserImpl > mxImpl;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

