/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: xmlfilterbase.cxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: sj $ $Date: 2007/09/04 17:06:16 $
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

#include "oox/core/xmlfilterbase.hxx"
#include <set>
#include <com/sun/star/document/XDocumentSubStorageSupplier.hpp>
#include <com/sun/star/embed/ElementModes.hpp>
#include <com/sun/star/embed/XTransactedObject.hpp>
#include <com/sun/star/xml/sax/InputSource.hpp>
#include <com/sun/star/xml/sax/XFastParser.hpp>
#include "oox/core/fasttokenhandler.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/relationshandler.hxx"
#include "oox/core/zipstorage.hxx"

using ::rtl::OUString;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::lang::XMultiServiceFactory;
using ::com::sun::star::embed::XStorage;
using ::com::sun::star::embed::XTransactedObject;
using ::com::sun::star::io::XInputStream;
using ::com::sun::star::io::XOutputStream;
using ::com::sun::star::document::XDocumentSubStorageSupplier;
using ::com::sun::star::xml::sax::XFastParser;
using ::com::sun::star::xml::sax::XFastTokenHandler;
using ::com::sun::star::xml::sax::XFastDocumentHandler;
using ::com::sun::star::xml::sax::InputSource;

namespace oox {
namespace core {

// ============================================================================

struct XmlFilterBaseImpl
{
    typedef ::std::map< OUString, RelationsRef > RelationsMap;

    OUString            maRels;
    OUString            maRelsSuffix;
    Reference< XFastTokenHandler >  mxTokenHandler;
    RelationsMap                    maRelationsMap;
    ::std::set< OUString >          maPictureSet;       /// Already copied picture stream names.
    Reference< XStorage >           mxPictureStorage;   /// Target model picture storage.

    explicit            XmlFilterBaseImpl();
};

// ----------------------------------------------------------------------------

XmlFilterBaseImpl::XmlFilterBaseImpl() :
    maRels( CREATE_OUSTRING( "_rels" ) ),
    maRelsSuffix( CREATE_OUSTRING( ".rels" ) ),
    mxTokenHandler( new FastTokenHandler )
{
}

// ============================================================================

XmlFilterBase::XmlFilterBase( const Reference< XMultiServiceFactory >& rxFactory ) :
    FilterBase( rxFactory ),
    mxImpl( new XmlFilterBaseImpl )
{
}

XmlFilterBase::~XmlFilterBase()
{
}

const oox::vml::DrawingPtr XmlFilterBase::getDrawings()
{
	oox::vml::DrawingPtr xRet;
	return xRet;
}


// ----------------------------------------------------------------------------

bool XmlFilterBase::importFragment( const Reference< XFastDocumentHandler >& xHandler, const OUString& rFragmentPath )
{
    InputSource aParserInput;
    aParserInput.sSystemId = rFragmentPath;
    aParserInput.aInputStream = openInputStream( rFragmentPath );

    try
    {
        Reference< XFastParser > xParser( getServiceFactory()->createInstance(
            CREATE_OUSTRING( "com.sun.star.xml.sax.FastParser" ) ), UNO_QUERY_THROW );
        xParser->setFastDocumentHandler( xHandler );
        xParser->setTokenHandler( mxImpl->mxTokenHandler );

        // register XML namespaces
        xParser->registerNamespace( CREATE_OUSTRING( "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), NMSP_EXCEL );
        xParser->registerNamespace( CREATE_OUSTRING( "urn:schemas-microsoft-com:office:office" ), NMSP_OFFICE );
        xParser->registerNamespace( CREATE_OUSTRING( "http://schemas.openxmlformats.org/presentationml/2006/main"), NMSP_PPT );
        xParser->registerNamespace( CREATE_OUSTRING( "urn:schemas-microsoft-com:office:word" ), NMSP_WORD );
        xParser->registerNamespace( CREATE_OUSTRING( "urn:schemas-microsoft-com:vml" ), NMSP_VML );
        xParser->registerNamespace( CREATE_OUSTRING( "http://schemas.openxmlformats.org/drawingml/2006/main" ), NMSP_DRAWINGML );
        xParser->registerNamespace( CREATE_OUSTRING( "http://schemas.openxmlformats.org/package/2006/relationships" ), NMSP_PACKAGE_RELATIONSHIPS );
        xParser->registerNamespace( CREATE_OUSTRING( "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ), NMSP_RELATIONSHIPS );
        xParser->registerNamespace( CREATE_OUSTRING( "http://www.w3.org/XML/1998/namespace" ), NMSP_XML );
        xParser->registerNamespace( CREATE_OUSTRING( "urn:schemas-microsoft-com:office:powerpoint" ), NMSP_POWERPOINT );
        xParser->registerNamespace( CREATE_OUSTRING( "urn:schemas-microsoft-com:office:activation" ), NMSP_ACTIVATION );

		// parse the stream
        xParser->parseStream( aParserInput );
    }
    catch( Exception& )
    {
        return false;
    }

    // success!
    return true;
}

RelationsRef XmlFilterBase::getRelations( const OUString& rRelPath )
{
    XmlFilterBaseImpl::RelationsMap::iterator aIter = mxImpl->maRelationsMap.find( rRelPath );
    if( aIter != mxImpl->maRelationsMap.end() )
        return aIter->second;

    RelationsRef xRelations( new Relations );

    sal_Int32 nIndex = rRelPath.lastIndexOf( '/' );

    OUString aRelsPath;
    if( nIndex != -1 )
        aRelsPath = rRelPath.copy( 0, nIndex + 1 );
    aRelsPath += mxImpl->maRels;

    if( nIndex != -1 )
    {
        aRelsPath += rRelPath.copy( nIndex );
    }
    else
    {
        aRelsPath += CREATE_OUSTRING( "/" );
        aRelsPath += rRelPath;
    }
    aRelsPath += mxImpl->maRelsSuffix;

    Reference< XFastDocumentHandler > xHandler( new RelationsFragmentHandler( aRelsPath, xRelations ) );
    importFragment( xHandler, aRelsPath );

    mxImpl->maRelationsMap[ rRelPath ] = xRelations;
    return xRelations;
}

OUString XmlFilterBase::copyPictureStream( const OUString& rPicturePath )
{
    // split source path into source storage path and stream name
    sal_Int32 nIndex = rPicturePath.lastIndexOf( sal_Unicode( '/' ) );
    OUString sPictureName;
    OUString sSourceStorageName;
    if( nIndex == -1 )
    {
        // root stream
        sPictureName = rPicturePath;
    }
    else
    {
        // sub stream
        sPictureName = rPicturePath.copy( nIndex + 1 );
        sSourceStorageName = rPicturePath.copy( 0, nIndex );
    }

    // check if we already copied this one!
    if( mxImpl->maPictureSet.find( rPicturePath ) == mxImpl->maPictureSet.end() ) try
    {
        // ok, not yet, copy stream to documents picture storage

        // first get the picture storage from our target model
        if( !mxImpl->mxPictureStorage.is() )
        {
            static const OUString sPictures = CREATE_OUSTRING( "Pictures" );
            Reference< XDocumentSubStorageSupplier > xDSSS( getModel(), UNO_QUERY_THROW );
            mxImpl->mxPictureStorage.set( xDSSS->getDocumentSubStorage(
                sPictures, ::com::sun::star::embed::ElementModes::WRITE ), UNO_QUERY_THROW );
        }

        StorageRef xSourceStorage = openSubStorage( sSourceStorageName, false );
        if( xSourceStorage.get() )
        {
            Reference< XStorage > xSourceXStorage = xSourceStorage->getXStorage();
            if( xSourceXStorage.is() )
            {
                xSourceXStorage->copyElementTo( sPictureName, mxImpl->mxPictureStorage, sPictureName );
                Reference< XTransactedObject > xTO( mxImpl->mxPictureStorage, UNO_QUERY_THROW );
                xTO->commit();
            }
        }
    }
    catch( Exception& )
    {
    }

    static const OUString sUrlPrefix = CREATE_OUSTRING( "vnd.sun.star.Package:Pictures/" );
    return sUrlPrefix + sPictureName;
}

StorageRef XmlFilterBase::implCreateStorage(
        Reference< XInputStream >& rxInStream, Reference< XOutputStream >& rxOutStream ) const
{
    StorageRef xStorage;
    if( rxInStream.is() )
        xStorage.reset( new ZipStorage( getServiceFactory(), rxInStream ) );
    else if( rxOutStream.is() )
        xStorage.reset( new ZipStorage( getServiceFactory(), rxOutStream ) );
    return xStorage;
}

// ============================================================================

} // namespace core
} // namespace oox

