/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: condformatbuffer.cxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 15:45:36 $
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

#include "oox/xls/condformatbuffer.hxx"
#include "oox/xls/tokenmapper.hxx"
#include "oox/xls/stylesbuffer.hxx"
#include "oox/xls/formulaparser.hxx"
#include "oox/core/propertyset.hxx"
#include "tokens.hxx"

#include <com/sun/star/sheet/ConditionOperator.hpp>
#include <com/sun/star/beans/PropertyValue.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include <com/sun/star/sheet/XSheetConditionalEntries.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheets.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/table/XCellRange.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/style/XStyleFamiliesSupplier.hpp>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/container/XNameContainer.hpp>


using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::XInterface;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::beans::PropertyValue;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::style::XStyleFamiliesSupplier;
using ::com::sun::star::container::XNameAccess;
using ::com::sun::star::container::XIndexAccess;
using ::com::sun::star::container::XNameContainer;
using ::com::sun::star::uno::Any;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::sheet::XSheetConditionalEntries;
using ::com::sun::star::sheet::XSpreadsheetDocument;
using ::com::sun::star::sheet::XSpreadsheets;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::table::XCellRange;
using ::com::sun::star::uno::Sequence;
using ::std::vector;
using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::oox::core::ContainerHelper;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

CondFormatDataBase::CondFormatDataBase()
{
    maFormulas.reserve(3);
}

CondFormatDataCellIs* CondFormatDataBase::getCellIs()
{
    return static_cast<CondFormatDataCellIs*>(this);
}

CondFormatDataText* CondFormatDataBase::getText()
{
    return static_cast<CondFormatDataText*>(this);
}

CondFormatDataTimePeriod* CondFormatDataBase::getTimePeriod()
{
    return static_cast<CondFormatDataTimePeriod*>(this);
}

// ----------------------------------------------------------------------------

CondFormatBuffer::CondFormatBuffer( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

void CondFormatBuffer::add( const CondFormatItemRef& aItem )
{
    maCondFormats.push_back(aItem);
}

void CondFormatBuffer::finalizeImport()
{
    try
    {
        Reference<XSpreadsheetDocument> xSpDoc( getDocument(), UNO_QUERY_THROW );
        Reference<XIndexAccess> xSheetsIA( xSpDoc->getSheets(), UNO_QUERY_THROW );

        vector<CondFormatItemRef>::const_iterator itr, itrEnd = maCondFormats.end();
        for ( itr = maCondFormats.begin(); itr != itrEnd; ++itr )
        {
            CondFormatItemRef pRef = *itr;
            const CellRangeAddress& rRange = pRef->maCellRange;
            Reference<XSpreadsheet> xSheet( xSheetsIA->getByIndex(pRef->mnSheetId), UNO_QUERY_THROW );
            Reference<XCellRange> xRange = xSheet->getCellRangeByPosition(
                rRange.StartColumn, rRange.StartRow, rRange.EndColumn, rRange.EndRow );

            PropertySet aProp( xRange );
            Reference<XSheetConditionalEntries> xEntries;
            aProp.getProperty( xEntries, CREATE_OUSTRING("ConditionalFormat") );
            vector<CondFormatDataRef>::const_iterator itr2, itr2End = pRef->maEntries.end();
            sal_Int32 nEntryId = 0;
            for ( itr2 = pRef->maEntries.begin(); itr2 != itr2End; ++itr2 )
                setEntry( xEntries, pRef, *itr2, nEntryId++ );

            aProp.setProperty( CREATE_OUSTRING("ConditionalFormat"), xEntries );
        }
    }
    catch ( const RuntimeException& )
    {
        OSL_ENSURE(false, "CondFormatBuffer::finalizeImport: runtime exception");
    }
}

void lclPushAny( vector<PropertyValue>& rPropList, const OUString& rName, const Any& rAny )
{
    rPropList.resize( rPropList.size() + 1 );
    rPropList.back().Name = rName;
    rPropList.back().Value = rAny;
}

void lclPushOperator( vector<PropertyValue>& rPropList,
                      ::com::sun::star::sheet::ConditionOperator eOp )
{
    lclPushAny(rPropList, CREATE_OUSTRING("Operator"), Any( eOp ) );
}

void lclPushFormulaString( vector<PropertyValue>& rPropList, const OUString& rFormula,
                           sal_Int32 nIndex )
{
    OUStringBuffer buf( CREATE_OUSTRING("Formula") );
    buf.append(nIndex);
    lclPushAny(rPropList, buf.makeStringAndClear(), Any( rFormula ) );
}

/** Do a best-effort conversion of Excel's time-period condition into
    something Calc can handle. */
void lclProcessTimePeriod( vector<PropertyValue>& rPropList, CondFormatDataTimePeriod* p )
{
    using namespace ::com::sun::star;

    Any any;
    switch ( p->mnTimePeriod )
    {
        case XML_last7Days:
        {
            lclPushOperator(rPropList, sheet::ConditionOperator_BETWEEN);
            lclPushFormulaString(rPropList, CREATE_OUSTRING("TODAY()-6"), 1);
            lclPushFormulaString(rPropList, CREATE_OUSTRING("TODAY()"),   2);
        }
        break;
        case XML_lastMonth:
        {
        }
        break;
        case XML_lastWeek:
        {
        }
        break;
        case XML_nextMonth:
        {
        }
        break;
        case XML_nextWeek:
        {
        }
        break;
        case XML_thisMonth:
        {
        }
        break;
        case XML_thisWeek:
        {
        }
        break;
        case XML_tomorrow:
        {
            lclPushOperator(rPropList, sheet::ConditionOperator_EQUAL);
            lclPushFormulaString(rPropList, CREATE_OUSTRING("TODAY()+1"), 1);
        }
        break;
        case XML_yesterday:
        {
            lclPushOperator(rPropList, sheet::ConditionOperator_EQUAL);
            lclPushFormulaString(rPropList, CREATE_OUSTRING("TODAY()-1"), 1);
        }
        break;
        case XML_today:
        {
            lclPushOperator(rPropList, sheet::ConditionOperator_EQUAL);
            lclPushFormulaString(rPropList, CREATE_OUSTRING("TODAY()"), 1);
        }
        break;
    }
}

void CondFormatBuffer::setEntry( const Reference<XSheetConditionalEntries>& rxEntries,
                                 const CondFormatItemRef& pRef,
                                 const CondFormatDataRef& pData,
                                 sal_Int32 nEntryId )
{
    using namespace ::com::sun::star;

    Any any;
    vector<PropertyValue> aPropList;
    aPropList.reserve(10);
    switch ( pData->mnType )
    {
        case XML_cellIs:
        {
            CondFormatDataCellIs* p = pData->getCellIs();

            sheet::ConditionOperator eOp = TokenMapper::convertOperator( p->mnOperator );
            if ( eOp == sheet::ConditionOperator_NONE )
                // Propabably unsupported operator type.
                return;

            any <<= eOp;
            aPropList.push_back( PropertyValue() );
            aPropList.back().Name = CREATE_OUSTRING("Operator");
            aPropList.back().Value = any;

            size_t nSize = p->maFormulas.size();
            if ( nSize >= 1 )
            {
                AnyFormulaContext aContext( any );
                getFormulaParser().importFormula( aContext, p->maFormulas[ 0 ] );
                lclPushAny(aPropList, CREATE_OUSTRING("Formula1"), any);
            }
            if ( nSize >= 2 )
            {
                AnyFormulaContext aContext( any );
                getFormulaParser().importFormula( aContext, p->maFormulas[ 1 ] );
                lclPushAny(aPropList, CREATE_OUSTRING("Formula2"), any);
            }
        }
        break;
        case XML_containsText:
        {
            lclPushOperator(aPropList, sheet::ConditionOperator_FORMULA);

            CondFormatDataText* p = pData->getText();
            if ( p->maFormulas.size() >= 1 )
            {
                AnyFormulaContext aContext( any );
                getFormulaParser().importFormula( aContext, p->maFormulas[ 0 ] );
                lclPushAny(aPropList, CREATE_OUSTRING("Formula1"), any);
            }
        }
        break;
        case XML_timePeriod:
            lclProcessTimePeriod( aPropList, pData->getTimePeriod() );
        break;
    }

    OUString aStyleName = registerStyleName( pRef, nEntryId, pData->mnDxfId );
    if ( !aStyleName.getLength() )
        return;

    any <<= aStyleName;
    aPropList.push_back( PropertyValue() );
    aPropList.back().Name = CREATE_OUSTRING("StyleName");
    aPropList.back().Value = any;

    size_t nSize = aPropList.size();
    Sequence<PropertyValue> aCond(nSize);
    for ( size_t i = 0; i < nSize; ++i )
        aCond[i] = aPropList[i];

    rxEntries->addNew(aCond);
}

OUString CondFormatBuffer::registerStyleName( const CondFormatItemRef& pRef, sal_Int32 nEntryId,
                                              sal_Int32 nDxfId )
{
    OUStringBuffer buf( CREATE_OUSTRING("Excel_CondFormat_") );
    buf.append( static_cast<sal_Int32>(pRef->mnSheetId+1) );
    buf.append( sal_Unicode( '_' ) );
    buf.append( static_cast<sal_Int32>(pRef->mnCondFormatId+1) );
    buf.append( sal_Unicode( '_' ) );
    buf.append( nEntryId+1 );

    OUString aStyleName = buf.makeStringAndClear();

    try
    {
        Reference<XMultiServiceFactory> xFactory( getDocument(), UNO_QUERY_THROW );
        Reference<XStyleFamiliesSupplier> xStyleSupplier( getDocument(), UNO_QUERY_THROW );
        Reference<XNameAccess> xSFamilies = xStyleSupplier->getStyleFamilies();
        Reference<XNameContainer> xNC( xSFamilies->getByName(CREATE_OUSTRING("CellStyles")), UNO_QUERY_THROW );
        Reference<XInterface> aStyle = xFactory->createInstance( CREATE_OUSTRING("com.sun.star.style.CellStyle") );
        Any any;
        any <<= aStyle;
        ContainerHelper::insertByUnusedName( xNC, any, aStyleName, sal_Unicode('_'), true );
        DxfRef aDxf = getStyles().getDxf(nDxfId);

        PropertySet aProp( aStyle );
        aDxf->getFont()->writeToPropertySet( aProp, FONT_PROPTYPE_CELL );
        aDxf->getFill()->writeToPropertySet( aProp );
        aDxf->getBorder()->writeToPropertySet( aProp );

        return aStyleName;
    }
    catch ( const RuntimeException& )
    {
        OSL_ENSURE(false, "OoxCondFormatContext::registerStyleName: failed to create a new cell style");
    }
    return OUString();
}

// ============================================================================

} // namespace xls
} // namespace oox
