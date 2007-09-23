/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffformatdetector.cxx,v $
 *
 *  $Revision: 1.1.2.11 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/16 16:24:46 $
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

#include "oox/xls/biffformatdetector.hxx"

#include <algorithm>
#include <rtl/strbuf.hxx>
#include <com/sun/star/io/XInputStream.hpp>
#include <comphelper/mediadescriptor.hxx>
#include "oox/core/binaryinputstream.hxx"
#include "oox/core/olestorage.hxx"

using ::rtl::OUString;
using ::rtl::OStringBuffer;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Sequence;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::RuntimeException;
using ::com::sun::star::uno::XInterface;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::beans::PropertyValue;
using ::com::sun::star::io::XInputStream;
using ::comphelper::MediaDescriptor;
using ::oox::core::BinaryInputStream;
using ::oox::core::OleStorage;
using ::oox::core::StorageRef;

namespace oox {
namespace xls {

// ============================================================================

Sequence< OUString > BiffFormatDetector_getSupportedServiceNames()
{
    Sequence< OUString > aServiceNames( 1 );
    aServiceNames[ 0 ] = CREATE_OUSTRING( "com.sun.star.frame.ExtendedTypeDetection" );
    return aServiceNames;
}

OUString BiffFormatDetector_getImplementationName()
{
    return CREATE_OUSTRING( "com.sun.star.comp.oox.BiffFormatDetector" );
}

Reference< XInterface > SAL_CALL BiffFormatDetector_createInstance( const Reference< XMultiServiceFactory >& rxFactory ) throw( Exception )
{
    return static_cast< ::cppu::OWeakObject* >( new BiffFormatDetector( rxFactory ) );
}

// ============================================================================

BiffFormatDetector::BiffFormatDetector( const Reference< XMultiServiceFactory >& rxFactory ) :
    mxFactory( rxFactory ),
    maWorkbookName( CREATE_OUSTRING( "Workbook" ) ),
    maBookName( CREATE_OUSTRING( "Book" ) )
{
}

BiffFormatDetector::~BiffFormatDetector()
{
}

BiffType BiffFormatDetector::detectStreamBiffVersion( BinaryInputStream& rInStream )
{
    BiffType eBiff = BIFF_UNKNOWN;
    if( rInStream.is() && rInStream.isSeekable() && (rInStream.getLength() > 4) )
    {
        sal_Int64 nOldPos = rInStream.tell();
        rInStream.seek( 0 );
        sal_uInt16 nBofId, nBofSize;
        rInStream >> nBofId >> nBofSize;

        if( (4 <= nBofSize) && (nBofSize <= 16) && (rInStream.tell() + nBofSize <= rInStream.getLength()) )
        {
            switch( nBofId )
            {
                case BIFF2_ID_BOF:
                    eBiff = BIFF2;
                break;
                case BIFF3_ID_BOF:
                    eBiff = BIFF3;
                break;
                case BIFF4_ID_BOF:
                    eBiff = BIFF4;
                break;
                case BIFF5_ID_BOF:
                {
                    if( 6 <= nBofSize )
                    {
                        sal_uInt16 nVersion;
                        rInStream >> nVersion;
                        // #i23425# #i44031# #i62752# there are some *really* broken documents out there...
                        switch( nVersion & 0xFF00 )
                        {
                            case 0:                 eBiff = BIFF5;  break;  // #i44031# #i62752#
                            case BIFF_BOF_BIFF2:    eBiff = BIFF2;  break;
                            case BIFF_BOF_BIFF3:    eBiff = BIFF3;  break;
                            case BIFF_BOF_BIFF4:    eBiff = BIFF4;  break;
                            case BIFF_BOF_BIFF5:    eBiff = BIFF5;  break;
                            case BIFF_BOF_BIFF8:    eBiff = BIFF8;  break;
                            default:    OSL_ENSURE( false,
                                OStringBuffer( "lclDetectStreamBiffVersion - unknown BIFF version: 0x" ).
                                append( static_cast< sal_Int32 >( nVersion ), 16 ).getStr() );
                        }
                    }
                }
                break;
                // else do nothing, no BIFF stream
            }
        }
        rInStream.seek( nOldPos );
    }
    return eBiff;
}

BiffType BiffFormatDetector::detectStorageBiffVersion( OUString& orWorkbookStreamName, StorageRef xStorage ) const
{
    BiffType eBiff = BIFF_UNKNOWN;
    if( xStorage.get() )
    {
        if( xStorage->isStorage() )
        {
            // try to open the "Book" stream
            BinaryInputStream aBookStrm5( xStorage->openInputStream( maBookName ) );
            BiffType eBookStrm5Biff = detectStreamBiffVersion( aBookStrm5 );

            // try to open the "Workbook" stream
            BinaryInputStream aBookStrm8( xStorage->openInputStream( maWorkbookName ) );
            BiffType eBookStrm8Biff = detectStreamBiffVersion( aBookStrm8 );

            // decide which stream to use
            if( (eBookStrm8Biff != BIFF_UNKNOWN) && ((eBookStrm5Biff == BIFF_UNKNOWN) || (eBookStrm8Biff > eBookStrm5Biff)) )
            {
                /*  Only "Workbook" stream exists; or both streams exist,
                    and "Workbook" has higher BIFF version than "Book" stream. */
                eBiff = eBookStrm8Biff;
                orWorkbookStreamName = maWorkbookName;
            }
            else if( eBookStrm5Biff != BIFF_UNKNOWN )
            {
                /*  Only "Book" stream exists; or both streams exist,
                    and "Book" has higher BIFF version than "Workbook" stream. */
                eBiff = eBookStrm5Biff;
                orWorkbookStreamName = maBookName;
            }
        }
        else
        {
            // no storage, try plain input stream from medium (even for BIFF5+)
            BinaryInputStream aStrm( xStorage->openInputStream( OUString() ) );
            eBiff = detectStreamBiffVersion( aStrm );
            orWorkbookStreamName = OUString();
        }
    }

    return eBiff;
}

// com.sun.star.lang.XServiceInfo interface -----------------------------------

OUString SAL_CALL BiffFormatDetector::getImplementationName() throw( RuntimeException )
{
    return BiffFormatDetector_getImplementationName();
}

sal_Bool SAL_CALL BiffFormatDetector::supportsService( const OUString& rService ) throw( RuntimeException )
{
    const Sequence< OUString > aServices( BiffFormatDetector_getSupportedServiceNames() );
    const OUString* pArray = aServices.getConstArray();
    const OUString* pArrayEnd = pArray + aServices.getLength();
    return ::std::find( pArray, pArrayEnd, rService ) != pArrayEnd;
}

Sequence< OUString > SAL_CALL BiffFormatDetector::getSupportedServiceNames() throw( RuntimeException )
{
    return BiffFormatDetector_getSupportedServiceNames();
}

// com.sun.star.document.XExtendedFilterDetect interface ----------------------

OUString SAL_CALL BiffFormatDetector::detect( Sequence< PropertyValue >& rDescriptor ) throw( RuntimeException )
{
    OUString aTypeName;

    MediaDescriptor aDescriptor( rDescriptor );
    aDescriptor.addInputStream();

    Reference< XInputStream > xInStrm( aDescriptor[ MediaDescriptor::PROP_INPUTSTREAM() ], UNO_QUERY );
    if( xInStrm.is() )
    {
        OUString aWorkbookName;
        StorageRef xStorage( new OleStorage( mxFactory, xInStrm, true ) );
        switch( detectStorageBiffVersion( aWorkbookName, xStorage ) )
        {
            case BIFF2:
            case BIFF3:
            case BIFF4: aTypeName = CREATE_OUSTRING( "calc_MS_Excel_40" );  break;
            case BIFF5: aTypeName = CREATE_OUSTRING( "calc_MS_Excel_95" );  break;
            case BIFF8: aTypeName = CREATE_OUSTRING( "calc_MS_Excel_97" );  break;
            default:;
        }
    }

    return aTypeName;
}

// ============================================================================

} // namespace xls
} // namespace oox

