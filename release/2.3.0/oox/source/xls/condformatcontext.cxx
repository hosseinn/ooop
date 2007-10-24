/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: condformatcontext.cxx,v $
 *
 *  $Revision: 1.1.2.17 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:20 $
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

#include "oox/xls/condformatcontext.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/tokenmapper.hxx"
#include "oox/xls/stylesbuffer.hxx"
#include "oox/core/propertyset.hxx"

#include <vector>

using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::xml::sax::SAXException;
using ::oox::core::AttributeList;
using ::oox::core::PropertySet;
using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::std::vector;

namespace oox {
namespace xls {

// ============================================================================

OoxCondFormatContext::OoxCondFormatContext( const WorksheetFragmentBase& rFragment ) :
    WorksheetContextBase( rFragment ),
    mnEntryId(0),
    mbValidRange(false)
{
}

void SAL_CALL OoxCondFormatContext::startDocument()
    throw (SAXException, RuntimeException)
{
}

void SAL_CALL OoxCondFormatContext::endDocument()
    throw (SAXException, RuntimeException)
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxCondFormatContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( conditionalFormatting ):
            return mbValidRange && (nElement == XLS_TOKEN( cfRule ));
        case XLS_TOKEN( cfRule ):
            return mbValidRange && (nElement == XLS_TOKEN( formula ));
    }
    return false;
}

void OoxCondFormatContext::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( conditionalFormatting ):
            importConditionalFormatting( rAttribs );
        break;
        case XLS_TOKEN( cfRule ):
            importCfRule( rAttribs );
        break;
        case XLS_TOKEN( formula ):
            importFormula( rAttribs );
        break;
    }
}

void OoxCondFormatContext::onEndElement( const OUString& rChars )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( conditionalFormatting ):
            if( mpCurItem.get() )
            {
                mpCurItem->mnSheetId = getSheetIndex();
                mpCurItem->mnCondFormatId = incCondFormatIndex();
                getCondFormats().add( mpCurItem );
            }
        break;
        case XLS_TOKEN( formula ):
            if( mpCurItem.get() && !mpCurItem->maEntries.empty() )
                mpCurItem->maEntries.back()->maFormulas.back() = rChars;
        break;
    }
}

void OoxCondFormatContext::importConditionalFormatting( const AttributeList& rAttribs )
{
    CellRangeAddress aRange;
    mbValidRange = getAddressConverter().convertToCellRange(
        aRange, rAttribs.getString( XML_sqref ), getSheetIndex(), true);

    if ( mbValidRange )
    {
        mpCurItem.reset(new CondFormatItem);
        mpCurItem->maCellRange = aRange;
    }
}

void OoxCondFormatContext::importCfRule( const AttributeList& rAttribs )
{
    // Possible rule type:
    // aboveAverage, beginsWith, cellIs, colorScale, containsBlanks, containsErrors,
    // containsText, dataBar, duplicateValues, endsWith, expression, iconSet,
    // notContainsBlanks, notContainsErrors, notContainsText, timePeriod, top10,
    // uniqueValues

    sal_Int32 nType = rAttribs.getToken( XML_type );
    CondFormatDataRef pData;
    switch ( nType )
    {
        case XML_cellIs:
        {
            pData.reset(new CondFormatDataCellIs);
            CondFormatDataCellIs* p = pData->getCellIs();
            p->mnOperator = rAttribs.getToken( XML_operator );
        }
        break;
        case XML_containsText:
        {
            pData.reset(new CondFormatDataText);
            CondFormatDataText* p = pData->getText();
            p->maText = rAttribs.getString( XML_text );
        }
        break;
        case XML_timePeriod:
        {
            pData.reset(new CondFormatDataTimePeriod);
            CondFormatDataTimePeriod* p = pData->getTimePeriod();
            p->mnTimePeriod = rAttribs.getToken( XML_timePeriod );
        }
        break;
        default:
            // unsupported conditional formatting type.
            return;
    }

    // Common attributes.
    pData->mnType     = nType;
    pData->mnPriority = rAttribs.getInteger( XML_priority, 1 );
    pData->mnDxfId    = rAttribs.getInteger( XML_dxfId, 0 );
    mpCurItem->maEntries.push_back(pData);
}

void OoxCondFormatContext::importFormula( const AttributeList& /*rAttribs*/ )
{
    if ( mpCurItem->maEntries.empty() )
        return;

    // Just add an empty new formula entry.
    mpCurItem->maEntries.back()->maFormulas.push_back( OUString() );
}

// ============================================================================

} // namespace xls
} // namespace oox

