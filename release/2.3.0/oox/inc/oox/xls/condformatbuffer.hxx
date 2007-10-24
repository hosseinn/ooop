/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: condformatbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 12:48:01 $
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

#ifndef OOX_XLS_CONDFORMATBUFFER_HXX
#define OOX_XLS_CONDFORMATBUFFER_HXX

#include "oox/xls/globaldatahelper.hxx"
#include "oox/core/containerhelper.hxx"
#include <rtl/ustrbuf.hxx>

#include <com/sun/star/table/CellRangeAddress.hpp>
#include <boost/shared_ptr.hpp>
#include <vector>

namespace com { namespace sun { namespace star {
    namespace sheet {
        class XSheetConditionalEntries;
    }
}}}

namespace oox {
namespace xls {

// ============================================================================

struct CondFormatDataCellIs;
struct CondFormatDataText;
struct CondFormatDataTimePeriod;

struct CondFormatDataBase
{
    sal_Int32 mnType;
    sal_Int32 mnPriority;
    sal_Int32 mnDxfId;

    ::std::vector< ::rtl::OUString > maFormulas;

    CondFormatDataCellIs* getCellIs();
    CondFormatDataText* getText();
    CondFormatDataTimePeriod* getTimePeriod();

    CondFormatDataBase();
};

struct CondFormatDataCellIs : public CondFormatDataBase
{
    sal_Int32 mnOperator;
};

struct CondFormatDataText : public CondFormatDataBase
{
    ::rtl::OUString maText;
};

struct CondFormatDataTimePeriod : public CondFormatDataBase
{
    sal_Int32 mnTimePeriod;
};

typedef ::boost::shared_ptr<CondFormatDataBase> CondFormatDataRef;

// ----------------------------------------------------------------------------

struct CondFormatItem
{
    sal_Int16       mnSheetId;
    sal_Int32       mnCondFormatId;
    ::com::sun::star::table::CellRangeAddress   maCellRange;
    ::std::vector<CondFormatDataRef>            maEntries;
};

typedef ::boost::shared_ptr<CondFormatItem> CondFormatItemRef;

// ----------------------------------------------------------------------------

class CondFormatBuffer : public GlobalDataHelper
{
public:
    explicit            CondFormatBuffer( const GlobalDataHelper& rGlobalData );

    void                add( const CondFormatItemRef& aItem );

    void                finalizeImport();

private:
    void                setEntry( const ::com::sun::star::uno::Reference<
                                      ::com::sun::star::sheet::XSheetConditionalEntries>& rxEntries,
                                  const CondFormatItemRef& pRef,
                                  const CondFormatDataRef& pData,
                                  sal_Int32 nEntryId );

    ::rtl::OUString     registerStyleName( const CondFormatItemRef& pRef, sal_Int32 nEntryId,
                                           sal_Int32 nDxfId );

private:
    ::std::vector<CondFormatItemRef>            maCondFormats;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

