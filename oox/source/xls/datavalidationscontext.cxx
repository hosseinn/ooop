/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: datavalidationscontext.cxx,v $
 *
 *  $Revision: 1.1.2.15 $
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

#include "oox/xls/datavalidationscontext.hxx"
#include "oox/core/propertyset.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/tokenmapper.hxx"
#include "oox/xls/formulaparser.hxx"
#include <com/sun/star/table/XCellRange.hpp>
#include <com/sun/star/sheet/ValidationType.hpp>
#include <com/sun/star/sheet/ValidationAlertStyle.hpp>
#include <com/sun/star/sheet/ConditionOperator.hpp>
#include <com/sun/star/sheet/XSheetCondition.hpp>
#include <com/sun/star/sheet/FormulaToken.hpp>
#include <com/sun/star/sheet/XFormulaTokens.hpp>
#include <com/sun/star/sheet/XMultiFormulaTokens.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>

#include <rtl/ustring.hxx>

#define DEBUG_OOX_DATA_VALIDATIONS 0

#if DEBUG_OOX_DATA_VALIDATIONS
#include <stdio.h>
#endif

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::com::sun::star::uno::Any;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::sheet::ConditionOperator;
using ::com::sun::star::sheet::ValidationType;
using ::com::sun::star::sheet::ValidationAlertStyle;
using ::com::sun::star::sheet::FormulaToken;
using ::com::sun::star::sheet::XFormulaTokens;
using ::com::sun::star::sheet::XMultiFormulaTokens;
using ::com::sun::star::sheet::XSheetCondition;
using ::com::sun::star::table::CellRangeAddress;
using ::com::sun::star::table::XCellRange;
using ::oox::core::AttributeList;
using ::oox::core::PropertySet;

namespace oox {
namespace xls {

// ============================================================================

/** Each instance of this struct represents each individual data validation
    item to import. */
struct DVItem
{
    struct Message
    {
        ::rtl::OUString maTitle;
        ::rtl::OUString maMessage;
    };

    CellRangeAddress maCellRange;

    ::std::auto_ptr< Message >  mpInputMsg;
    ::std::auto_ptr< Message >  mpErrorMsg;

    ::rtl::OUString         maFormula1;
    ::rtl::OUString         maFormula2;

    ValidationType          maType;
    ValidationAlertStyle    maAlertStyle;
    ConditionOperator       maOperator;

    bool            mbAllowBlank;
    bool            mbShowDropDown;

    DVItem();
};

DVItem::DVItem() :
    mpInputMsg(NULL),
    mpErrorMsg(NULL),
    maType(::com::sun::star::sheet::ValidationType_ANY),
    maAlertStyle(::com::sun::star::sheet::ValidationAlertStyle_INFO),
    maOperator(::com::sun::star::sheet::ConditionOperator_NONE),
    mbAllowBlank(true),
    mbShowDropDown(true)
{
}

// ============================================================================

OoxDataValidationsContext::OoxDataValidationsContext( const WorksheetFragmentBase& rFragment ) :
    WorksheetContextBase( rFragment )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxDataValidationsContext::onCanCreateContext( sal_Int32 nElement )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( dataValidations ):
            return (nElement == XLS_TOKEN( dataValidation ));
        case XLS_TOKEN( dataValidation ):
            return (nElement == XLS_TOKEN( formula1 )) ||
                   (nElement == XLS_TOKEN( formula2 ));
    }
    return false;
}

void OoxDataValidationsContext::onStartElement( const AttributeList& rAttribs )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( dataValidations ):
            importDataValidations( rAttribs );
        break;
        case XLS_TOKEN( dataValidation ):
            importDataValidation( rAttribs );
        break;
    }
}

void OoxDataValidationsContext::onEndElement( const OUString& rChars )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( dataValidation ):
            setDataValidation();
        break;
        case XLS_TOKEN( formula1 ):
            if( mpCurItem.get() ) mpCurItem->maFormula1 = rChars;
        break;
        case XLS_TOKEN( formula2 ):
            if( mpCurItem.get() ) mpCurItem->maFormula2 = rChars;
        break;
    }
}

void OoxDataValidationsContext::importDataValidations( const AttributeList& /*rAttribs*/ )
{
    // ----------------------------------------
    // count (unsigned int)
    // disablePrompts (bool)
    // xWindow (unsigned int)
    // yWindow (unsigned int)
    // ----------------------------------------
}

ValidationType lclTranslateType( const sal_Int32 nToken )
{
    using namespace ::com::sun::star::sheet;

    switch ( nToken )
    {
        case XML_custom:        return ValidationType_CUSTOM;
        case XML_date:          return ValidationType_DATE;
        case XML_decimal:       return ValidationType_DECIMAL;
        case XML_list:          return ValidationType_LIST;
        case XML_none:          return ValidationType_ANY;
        case XML_textLength:    return ValidationType_TEXT_LEN;
        case XML_time:          return ValidationType_TIME;
        case XML_whole:         return ValidationType_WHOLE;
    }

    return ValidationType_ANY;
}

ValidationAlertStyle lclTranslateAlertStyle( const sal_Int32 nToken )
{
    using namespace ::com::sun::star::sheet;

    switch ( nToken )
    {
        case XML_information:   return ValidationAlertStyle_INFO;
        case XML_stop:          return ValidationAlertStyle_STOP;
        case XML_warning:       return ValidationAlertStyle_WARNING;
    }
    return ValidationAlertStyle_STOP;
}

ConditionOperator lclTranslateOp( const sal_Int32 nToken )
{
    using namespace ::com::sun::star::sheet;

    switch ( nToken )
    {
        case XML_between:               return ConditionOperator_BETWEEN;
        case XML_equal:                 return ConditionOperator_EQUAL;
        case XML_greaterThan:           return ConditionOperator_GREATER;
        case XML_greaterThanOrEqual:    return ConditionOperator_GREATER_EQUAL;
        case XML_lessThan:              return ConditionOperator_LESS;
        case XML_lessThanOrEqual:       return ConditionOperator_LESS_EQUAL;
        case XML_notBetween:            return ConditionOperator_NOT_BETWEEN;
        case XML_notEqual:              return ConditionOperator_NOT_EQUAL;
    }

    return ConditionOperator_NONE;
}

#if DEBUG_OOX_DATA_VALIDATIONS
void debugOp( ConditionOperator aOp )
{
    using namespace ::com::sun::star::sheet;
    switch ( aOp )
    {
        case ConditionOperator_BETWEEN:
            printf("BETWEEN\n");
            break;
        case ConditionOperator_EQUAL:
            printf("EQUAL\n");
            break;
        case ConditionOperator_GREATER:
            printf("GREATER\n");
            break;
        case ConditionOperator_GREATER_EQUAL:
            printf("GREATER_EQUAL\n");
            break;
        case ConditionOperator_LESS:
            printf("LESS\n");
            break;
        case ConditionOperator_LESS_EQUAL:
            printf("LESS_EQUAL\n");
            break;
        case ConditionOperator_NOT_BETWEEN:
            printf("NOT_BETWEEN\n");
            break;
        case ConditionOperator_NOT_EQUAL:
            printf("NOT_EQUAL\n");
            break;
        case ConditionOperator_NONE:
            printf("NONE\n");
            break;
    }
}
#endif

void OoxDataValidationsContext::importDataValidation( const AttributeList& rAttribs )
{
    // ----------------------------------------
    // allowBlank (bool)
    // error (string)
    // errorStyle (information, stop, warning)
    // errorTitle (string)
    // imeMode (enum - see 3.18.20)
    // operator (between, equal, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, notBetween, notEqual)
    // prompt (string)
    // promptTitle (string)
    // showDropDown (bool)
    // showErrorMessage (bool)
    // showInputMessage (bool)
    // sqref
    // type (custom, date, decimal, list, none, textLength, time, whole)
    // ----------------------------------------

    // Remove any previous data validation item before we start a new one.
    mpCurItem.reset(NULL);

    OUString aRef = rAttribs.getString( XML_sqref );
    if ( !aRef.getLength() )
        // No reference, no data validation!
        return;

    CellRangeAddress aAddr;
    bool bValid = getAddressConverter().convertToCellRange( aAddr, aRef, getSheetIndex(), true );
    if ( !bValid )
        // No valid reference, no data validation!
        return;

    mpCurItem.reset( new DVItem );
    mpCurItem->maCellRange = aAddr;

    mpCurItem->maType       = lclTranslateType( rAttribs.getToken( XML_type, XML_none ) );
    mpCurItem->maOperator   = TokenMapper::convertOperator( rAttribs.getToken( XML_operator, XML_between ) );
    mpCurItem->maAlertStyle = lclTranslateAlertStyle( rAttribs.getToken( XML_errorStyle, XML_stop ) );

    mpCurItem->mbAllowBlank   = rAttribs.getBool( XML_allowBlank, false );
    mpCurItem->mbShowDropDown = rAttribs.getBool( XML_showDropDown, false );

    if ( rAttribs.getBool( XML_showInputMessage, false ) )
    {
        mpCurItem->mpInputMsg.reset( new DVItem::Message );
        mpCurItem->mpInputMsg->maTitle = rAttribs.getString( XML_promptTitle );
        mpCurItem->mpInputMsg->maMessage = rAttribs.getString( XML_prompt );
    }

    if ( rAttribs.getBool( XML_showErrorMessage, false ) )
    {
        mpCurItem->mpErrorMsg.reset( new DVItem::Message );
        mpCurItem->mpErrorMsg->maTitle = rAttribs.getString( XML_errorTitle );
        mpCurItem->mpErrorMsg->maMessage = rAttribs.getString( XML_error );
    }
}

void OoxDataValidationsContext::setDataValidation()
{
    if ( !mpCurItem.get() )
        return;

    const OUString aValidName = CREATE_OUSTRING("Validation");

    Reference< XCellRange > xCellRange = getWorksheetHelper().getCellRange( mpCurItem->maCellRange );
    PropertySet aPropSet( xCellRange );
    Any aValidation;
    aPropSet.getProperty( aValidation, aValidName );
    PropertySet aValidPropSet( aValidation, false ); // It doesn't support XMultiPropertySet.

    aValidPropSet.setProperty( CREATE_OUSTRING("Type"), mpCurItem->maType );
    aValidPropSet.setProperty( CREATE_OUSTRING("IgnoreBlankCells"), mpCurItem->mbAllowBlank );
    aValidPropSet.setProperty( CREATE_OUSTRING("ShowList"), mpCurItem->mbShowDropDown );
    aValidPropSet.setProperty( CREATE_OUSTRING("ErrorAlertStyle"), mpCurItem->maAlertStyle );

    // Check input message.
    aValidPropSet.setProperty( CREATE_OUSTRING("ShowInputMessage"), mpCurItem->mpInputMsg.get() != NULL );
    if ( mpCurItem->mpInputMsg.get() )
    {
        aValidPropSet.setProperty( CREATE_OUSTRING("InputTitle"),   mpCurItem->mpInputMsg->maTitle );
        aValidPropSet.setProperty( CREATE_OUSTRING("InputMessage"), mpCurItem->mpInputMsg->maMessage );
    }

    // Check error message.
    aValidPropSet.setProperty( CREATE_OUSTRING("ShowErrorMessage"), mpCurItem->mpErrorMsg.get() != NULL );
    if ( mpCurItem->mpErrorMsg.get() )
    {
        aValidPropSet.setProperty( CREATE_OUSTRING("ErrorTitle"),   mpCurItem->mpErrorMsg->maTitle );
        aValidPropSet.setProperty( CREATE_OUSTRING("ErrorMessage"), mpCurItem->mpErrorMsg->maMessage );
    }

    // Set formulas and the operator.
    Reference< XSheetCondition > xSheetCond( aValidation, UNO_QUERY );
    Reference< XMultiFormulaTokens > xMultiTokens( aValidation, UNO_QUERY );

    if ( xSheetCond.is() && xMultiTokens.is() )
    {
        xSheetCond->setOperator( mpCurItem->maOperator );

        if ( mpCurItem->maFormula1.getLength() && xMultiTokens->getCount() >= 1 )
        {
            MultiFormulaContext aContext( xMultiTokens, 0 );
            getFormulaParser().importFormula( aContext, mpCurItem->maFormula1 );
        }

        if ( mpCurItem->maFormula2.getLength() && xMultiTokens->getCount() >= 2 )
        {
            MultiFormulaContext aContext( xMultiTokens, 1 );
            getFormulaParser().importFormula( aContext, mpCurItem->maFormula2 );
        }
    }

    // Commit the change.
    aPropSet.setProperty( aValidName, aValidation );
}

} // namespace xls
} // namespace oox

