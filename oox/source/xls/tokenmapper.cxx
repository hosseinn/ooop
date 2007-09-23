/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: tokenmapper.cxx,v $
 *
 *  $Revision: 1.1.2.1 $
 *
 *  last change: $Author: kohei $ $Date: 2007/07/23 22:21:12 $
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

#include "oox/xls/tokenmapper.hxx"
#include "tokens.hxx"

using ::com::sun::star::sheet::ConditionOperator;

namespace oox {
namespace xls {

// ============================================================================

TokenMapper::TokenMapper( const GlobalDataHelper& rGlobalData ) :
    GlobalDataHelper( rGlobalData )
{
}

ConditionOperator TokenMapper::convertOperator( const sal_Int32 nToken )
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

// ============================================================================

} // namespace xls
} // namespace oox

