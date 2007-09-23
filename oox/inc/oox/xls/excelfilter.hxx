/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: excelfilter.hxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/12 14:15:53 $
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

#ifndef OOX_XLS_EXCELFILTER_HXX
#define OOX_XLS_EXCELFILTER_HXX

#include "oox/core/xmlfilterbase.hxx"
#include "oox/core/binaryfilterbase.hxx"
#include "oox/xls/globaldatahelper.hxx"

namespace oox {
namespace xls {

// ============================================================================

class ExcelFilter : public ::oox::core::XmlFilterBase
{
public:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiServiceFactory > XMultiServiceFactoryRef;

public:
    explicit            ExcelFilter( const XMultiServiceFactoryRef& rxFactory );
    virtual             ~ExcelFilter();

    virtual bool        importDocument() throw();
    virtual bool        exportDocument() throw();

    virtual sal_Int32   getSchemeClr( sal_Int32 nColorSchemeToken );

private:
    virtual ::rtl::OUString implGetImplementationName() const;

private:
    GlobalDataRef       mxGlobalData;
};

// ============================================================================

class ExcelBiffFilter : public ::oox::core::BinaryFilterBase
{
public:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiServiceFactory > XMultiServiceFactoryRef;

public:
    explicit            ExcelBiffFilter( const XMultiServiceFactoryRef& rxFactory );
    virtual             ~ExcelBiffFilter();

    virtual bool        importDocument() throw();
    virtual bool        exportDocument() throw();

private:
    virtual ::rtl::OUString implGetImplementationName() const;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

