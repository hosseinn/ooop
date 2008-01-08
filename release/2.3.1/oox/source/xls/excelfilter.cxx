/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: excelfilter.cxx,v $
 *
 *  $Revision: 1.1.2.25 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/29 14:01:34 $
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

#include "oox/xls/excelfilter.hxx"
#include "oox/core/binaryinputstream.hxx"
#include "oox/xls/biffformatdetector.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/themebuffer.hxx"
#include "oox/xls/workbookfragment.hxx"
#include "oox/dump/biffdumper.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::XInterface;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::xml::sax::XFastDocumentHandler;
using ::oox::core::BinaryFilterBase;
using ::oox::core::BinaryInputStream;
using ::oox::core::Relation;
using ::oox::core::Relations;
using ::oox::core::XmlFilterBase;

namespace oox {
namespace xls {

// ============================================================================

OUString SAL_CALL ExcelFilter_getImplementationName() throw()
{
    return CREATE_OUSTRING( "com.sun.star.comp.oox.ExcelFilter" );
}

Sequence< OUString > SAL_CALL ExcelFilter_getSupportedServiceNames() throw()
{
    OUString aServiceName = CREATE_OUSTRING( "com.sun.star.comp.oox.ExcelFilter" );
    Sequence< OUString > aSeq( &aServiceName, 1 );
    return aSeq;
}

Reference< XInterface > SAL_CALL ExcelFilter_createInstance(
        const Reference< XMultiServiceFactory >& rxFactory ) throw( Exception )
{
    return static_cast< ::cppu::OWeakObject* >( new ExcelFilter( rxFactory ) );
}

// ----------------------------------------------------------------------------

ExcelFilter::ExcelFilter( const Reference< XMultiServiceFactory >& rxFactory ) :
    XmlFilterBase( rxFactory )
{
}

ExcelFilter::~ExcelFilter()
{
}

bool ExcelFilter::importDocument() throw()
{
    bool bRet = false;
    if( const Relations* pRelations = getRelations( OUString() ).get() )
    {
        if( const Relation* pDocRel = pRelations->getRelationByType( CREATE_RELATIONS_TYPE( "officeDocument" ) ).get() )
        {
            mxGlobalData = GlobalDataHelper::createGlobalDataStruct( this );
            if( mxGlobalData.get() )
            {
                // initial construction of global data object
                GlobalDataHelper aGlobalDataHelper( *mxGlobalData );
                // create workbook fragment handler and import the document
                Reference< XFastDocumentHandler > xHandler = new OoxWorkbookFragment( aGlobalDataHelper, pDocRel->msTarget );
                bRet = importFragment( xHandler, pDocRel->msTarget );
            }
            // must reset global data here, to prevent cyclic referencing
            mxGlobalData.reset();
        }
    }
    return bRet;
}

bool ExcelFilter::exportDocument() throw()
{
    return false;
}

sal_Int32 ExcelFilter::getSchemeClr( sal_Int32 nColorSchemeToken )
{
    return GlobalDataHelper( *mxGlobalData ).getTheme().getColorByToken( nColorSchemeToken );
}

OUString ExcelFilter::implGetImplementationName() const
{
    return ExcelFilter_getImplementationName();
}

// ============================================================================

OUString SAL_CALL ExcelBiffFilter_getImplementationName() throw()
{
    return CREATE_OUSTRING( "com.sun.star.comp.oox.ExcelBiffFilter" );
}

Sequence< OUString > SAL_CALL ExcelBiffFilter_getSupportedServiceNames() throw()
{
    OUString aServiceName = CREATE_OUSTRING( "com.sun.star.comp.oox.ExcelBiffFilter" );
    Sequence< OUString > aSeq( &aServiceName, 1 );
    return aSeq;
}

Reference< XInterface > SAL_CALL ExcelBiffFilter_createInstance(
        const Reference< XMultiServiceFactory >& rxFactory ) throw( Exception )
{
    return static_cast< ::cppu::OWeakObject* >( new ExcelBiffFilter( rxFactory ) );
}

// ----------------------------------------------------------------------------

ExcelBiffFilter::ExcelBiffFilter( const Reference< XMultiServiceFactory >& rxFactory ) :
    BinaryFilterBase( rxFactory )
{
}

ExcelBiffFilter::~ExcelBiffFilter()
{
}

bool ExcelBiffFilter::importDocument() throw()
{
#if OOX_INCLUDE_DUMPER
    {
        ::oox::dump::biff::Dumper aDumper( *this );
        aDumper.dump();
        if( !aDumper.isImportEnabled() )
            return aDumper.isValid();
    }
#endif

    bool bRet = false;

    // detect BIFF version and workbook stream name
    BiffFormatDetector aDetector( getServiceFactory() );
    OUString aWorkbookName;
    BiffType eBiff = aDetector.detectStorageBiffVersion( aWorkbookName, getStorage() );
    BinaryInputStream aInStrm( getStorage()->openInputStream( aWorkbookName ) );
    OSL_ENSURE( (eBiff != BIFF_UNKNOWN) && aInStrm.is(), "ExcelBiffFilter::ExcelBiffFilter - invalid file format" );

    if( (eBiff != BIFF_UNKNOWN) && aInStrm.is() )
    {
        GlobalDataRef xGlobalData = GlobalDataHelper::createGlobalDataStruct( this, eBiff );
        if( xGlobalData.get() )
        {
            BiffInputStream aBiffStream( aInStrm );
            GlobalDataHelper aGlobalDataHelper( *xGlobalData );
            BiffWorkbookFragment aFragment( aGlobalDataHelper );
            bRet = aFragment.importWorkbook( aBiffStream );
        }
    }
    return bRet;
}

bool ExcelBiffFilter::exportDocument() throw()
{
    return false;
}

OUString ExcelBiffFilter::implGetImplementationName() const
{
    return ExcelBiffFilter_getImplementationName();
}

// ============================================================================

} // namespace xls
} // namespace oox

