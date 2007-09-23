/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: presentationfragmenthandler.cxx,v $
 *
 *  $Revision: 1.1.2.19 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/24 15:36:32 $
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

#include "comphelper/anytostring.hxx"
#include "cppuhelper/exc_hlp.hxx"

#include <com/sun/star/drawing/XMasterPagesSupplier.hpp>
#include <com/sun/star/drawing/XDrawPages.hpp>
#include <com/sun/star/drawing/XDrawPagesSupplier.hpp>
#include <com/sun/star/drawing/XMasterPageTarget.hpp>
#include <com/sun/star/style/XStyleFamiliesSupplier.hpp>
#include <com/sun/star/style/XStyle.hpp>
#include <com/sun/star/presentation/XPresentationPage.hpp>

#include "oox/drawingml/theme.hxx"
#include "oox/drawingml/drawingmltypes.hxx"
#include "oox/drawingml/themefragmenthandler.hxx"
#include "oox/drawingml/textliststylecontext.hxx"
#include "oox/ppt/pptshape.hxx"
#include "oox/ppt/presentationfragmenthandler.hxx"
#include "oox/ppt/slidefragmenthandler.hxx"
#include "oox/ppt/layoutfragmenthandler.hxx"
#include "oox/ppt/pptimport.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::presentation;
using namespace ::com::sun::star::xml::sax;

#define C2U(x) OUString( RTL_CONSTASCII_USTRINGPARAM( x ) )

namespace oox { namespace ppt {

PresentationFragmentHandler::PresentationFragmentHandler( const XmlFilterRef& xFilter, const OUString& rFragmentPath ) throw()
: FragmentHandler( xFilter, rFragmentPath )
, mpTextListStyle( new TextListStyle())
{
}

PresentationFragmentHandler::~PresentationFragmentHandler() throw()
{

}
void PresentationFragmentHandler::startDocument() throw (SAXException, RuntimeException)
{
}

void PresentationFragmentHandler::endDocument() throw (SAXException, RuntimeException)
{
	try
	{
        Reference< frame::XModel > xModel( mxFilter->getModel() );
		Reference< drawing::XDrawPage > xSlide;
		sal_uInt32 nSlide;

		// importing slide pages and its corresponding notes page
		Reference< drawing::XDrawPagesSupplier > xDPS( xModel, uno::UNO_QUERY_THROW );
		Reference< drawing::XDrawPages > xDrawPages( xDPS->getDrawPages(), uno::UNO_QUERY_THROW );

		for( nSlide = 0; nSlide < maSlidesVector.size(); nSlide++ )
		{
			if( nSlide == 0 )
				xDrawPages->getByIndex( 0 ) >>= xSlide;
			else
				xSlide = xDrawPages->insertNewByIndex( nSlide );

			OUString aSlideFragmentPath;
			oox::core::RelationPtr pRelation( mxRelations->getRelationById( maSlidesVector[ nSlide ] ) );
			if ( pRelation )
			{
				SlidePersistPtr pMasterPersistPtr;
				SlidePersistPtr pSlidePersistPtr( new SlidePersist( sal_False, sal_False, xSlide,
									ShapePtr( new PPTShape( Slide, "com.sun.star.drawing.GroupShape" ) ), mpTextListStyle ) );

				aSlideFragmentPath = resolveRelativePath( pRelation->msTarget );
                SlideFragmentHandler* pSlideFragmentHandler( new SlideFragmentHandler( mxFilter, aSlideFragmentPath, pSlidePersistPtr, Slide ) );
				Reference< XFastDocumentHandler > xSlideFragmentHandler( pSlideFragmentHandler );

				// importing the corresponding masterpage/layout
				RelationPtr pLayout( pSlideFragmentHandler->getRelations()->
					getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" ) ) );
				if ( pLayout )
				{
					// importing layout
					OUString aLayoutFragmentPath( resolveRelativePath( aSlideFragmentPath, pLayout->msTarget ) );
					if ( aLayoutFragmentPath.getLength() )
					{
						RelationPtr pMaster( mxFilter->getRelations( aLayoutFragmentPath )->
							getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" ) ) );
						if( pMaster )
						{
							OUString aMasterFragmentPath( resolveRelativePath( aLayoutFragmentPath, pMaster->msTarget ) );
							if ( aMasterFragmentPath.getLength() )
							{
								// check if the corresponding masterpage+layout has already been imported
                                std::vector< SlidePersistPtr >& rMasterPages( (dynamic_cast< PowerPointImport& >( *getFilter() )).getMasterPages() );
								std::vector< SlidePersistPtr >::iterator aIter( rMasterPages.begin() );
								while( aIter != rMasterPages.end() )
								{
									if ( ( (*aIter)->getPath() == aMasterFragmentPath ) && ( (*aIter)->getLayoutPath() == aLayoutFragmentPath ) )
									{
										pMasterPersistPtr = *aIter;
										break;
									}
									aIter++;
								}
								if ( aIter == rMasterPages.end() )
								{	// masterpersist not found, we have to load it
									Reference< drawing::XDrawPage > xMasterPage;
									Reference< drawing::XMasterPagesSupplier > xMPS( xModel, uno::UNO_QUERY_THROW );
									Reference< drawing::XDrawPages > xMasterPages( xMPS->getMasterPages(), uno::UNO_QUERY_THROW );

                                    if( !((dynamic_cast< PowerPointImport& >( *getFilter() )).getMasterPages().size() ))
										xMasterPages->getByIndex( 0 ) >>= xMasterPage;
									else
										xMasterPage = xMasterPages->insertNewByIndex( xMasterPages->getCount() );

									pMasterPersistPtr = SlidePersistPtr( new SlidePersist( sal_True, sal_False, xMasterPage,
										ShapePtr( new PPTShape( Master, "com.sun.star.drawing.GroupShape" ) ), mpTextListStyle ) );
									pMasterPersistPtr->setLayoutPath( aLayoutFragmentPath );
                                    (dynamic_cast< PowerPointImport& >( *getFilter() )).getMasterPages().push_back( pMasterPersistPtr );
                                    (dynamic_cast< PowerPointImport& >( *getFilter() )).setActualSlidePersist( pMasterPersistPtr );
                                    SlideFragmentHandler* pMasterFragmentHandler( new SlideFragmentHandler( mxFilter, aMasterFragmentPath, pMasterPersistPtr, Master ) );
									Reference< XFastDocumentHandler > xMasterFragmentHandler( pMasterFragmentHandler );

									// set the correct theme
									RelationPtr pThemeRelation( pMasterFragmentHandler->getRelations()->
										getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" ) ) );
									if ( pThemeRelation )
									{
										OUString aThemeFragmentPath( resolveRelativePath( aMasterFragmentPath, pThemeRelation->msTarget ) );
                                        std::map< OUString, oox::drawingml::ThemePtr >& rThemes( (dynamic_cast< PowerPointImport& >( *getFilter() )).getThemes() );
                                        std::map< OUString, oox::drawingml::ThemePtr >::iterator aIter2( rThemes.find( aThemeFragmentPath ) );
                                        if( aIter2 == rThemes.end() )
										{
											oox::drawingml::ThemePtr pThemePtr( new oox::drawingml::Theme() );
											pMasterPersistPtr->setTheme( pThemePtr );
											Reference< XFastDocumentHandler > xThemeFragmentHandler(
                                                new ThemeFragmentHandler( mxFilter, aThemeFragmentPath, pThemePtr ) );
                                            mxFilter->importFragment( xThemeFragmentHandler, aThemeFragmentPath );
											rThemes[ aThemeFragmentPath ] = pThemePtr;
										}
										else
										{
                                            pMasterPersistPtr->setTheme( (*aIter2).second );
										}
									}
									Reference< XFastDocumentHandler > xLayoutFragmentHandler(
										new LayoutFragmentHandler( mxFilter, aLayoutFragmentPath, pMasterPersistPtr ) );
									importSlide( xMasterFragmentHandler, aMasterFragmentPath, pMasterPersistPtr );
                                    mxFilter->importFragment( xLayoutFragmentHandler, aLayoutFragmentPath );
									pMasterPersistPtr->createXShapes( xModel );
								}
							}
						}
					}
				}

				// importing slide page
				pSlidePersistPtr->setMasterPersist( pMasterPersistPtr );
				pSlidePersistPtr->setTheme( pMasterPersistPtr->getTheme() );
				Reference< drawing::XMasterPageTarget > xMasterPageTarget( pSlidePersistPtr->getPage(), UNO_QUERY );
				if( xMasterPageTarget.is() )
					xMasterPageTarget->setMasterPage( pMasterPersistPtr->getPage() );
                (dynamic_cast< PowerPointImport& >( *getFilter() )).getDrawPages().push_back( pSlidePersistPtr );
                (dynamic_cast< PowerPointImport& >( *getFilter() )).setActualSlidePersist( pSlidePersistPtr );
				importSlide( xSlideFragmentHandler, aSlideFragmentPath, pSlidePersistPtr );
				pSlidePersistPtr->createXShapes( xModel );

				// now importing the notes page
				RelationPtr pNotesRelation( pSlideFragmentHandler->getRelations()->
					getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" ) ) );
				if ( pNotesRelation )
				{
					OUString aNotesFragmentPath( resolveRelativePath( aSlideFragmentPath, pNotesRelation->msTarget ) );
					Reference< XPresentationPage > xPresentationPage( xSlide, UNO_QUERY );
					if ( xPresentationPage.is() )
					{
						Reference< XDrawPage > xNotesPage( xPresentationPage->getNotesPage() );
						if ( xNotesPage.is() )
						{
							SlidePersistPtr pNotesPersistPtr( new SlidePersist( sal_False, sal_True, xNotesPage,
								ShapePtr( new PPTShape( Slide, "com.sun.star.drawing.GroupShape" ) ), mpTextListStyle ) );
                            Reference< XFastDocumentHandler > xNotesFragmentHandler( new SlideFragmentHandler( mxFilter, aNotesFragmentPath, pNotesPersistPtr, Slide ) );
                            (dynamic_cast< PowerPointImport& >( *getFilter() )).getNotesPages().push_back( pNotesPersistPtr );
                            (dynamic_cast< PowerPointImport& >( *getFilter() )).setActualSlidePersist( pNotesPersistPtr );
							importSlide( xNotesFragmentHandler, aNotesFragmentPath, pNotesPersistPtr );
							pNotesPersistPtr->createXShapes( xModel );
						}
					}
				}
			}
		}
	}
	catch( uno::Exception& )
	{
        OSL_ENSURE( false,
			(rtl::OString("oox::ppt::PresentationFragmentHandler::EndDocument(), "
					"exception caught: ") +
			rtl::OUStringToOString(
				comphelper::anyToString( cppu::getCaughtException() ),
				RTL_TEXTENCODING_UTF8 )).getStr() );

	}

	// todo	error handling;
}

// CT_Presentation
Reference< XFastContextHandler > PresentationFragmentHandler::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
	case NMSP_PPT|XML_presentation:
	case NMSP_PPT|XML_sldMasterIdLst:
	case NMSP_PPT|XML_notesMasterIdLst:
	case NMSP_PPT|XML_sldIdLst:
		break;
	case NMSP_PPT|XML_sldMasterId:
		maSlideMasterVector.push_back( xAttribs->getOptionalValue( NMSP_RELATIONSHIPS|XML_id ) );
		break;
	case NMSP_PPT|XML_sldId:
		maSlidesVector.push_back( xAttribs->getOptionalValue( NMSP_RELATIONSHIPS|XML_id ) );
		break;
	case NMSP_PPT|XML_notesMasterId:
		maNotesMasterVector.push_back( xAttribs->getOptionalValue(NMSP_RELATIONSHIPS|XML_id ) );
		break;
	case NMSP_PPT|XML_sldSz:
		maSlideSize = GetSize2D( xAttribs );
		break;
	case NMSP_PPT|XML_notesSz:
		maNotesSize = GetSize2D( xAttribs );
		break;
	case NMSP_PPT|XML_defaultTextStyle:
		xRet.set( new TextListStyleContext( this, mpTextListStyle ) );
		break;
	}
	if ( !xRet.is() )
		xRet.set( this );
	return xRet;
}

bool PresentationFragmentHandler::importSlide( Reference< XFastDocumentHandler >& rxSlideFragmentHandler,
		const rtl::OUString& rSlideFragmentPath, const SlidePersistPtr pSlidePersistPtr )
{
	Reference< drawing::XDrawPage > xSlide( pSlidePersistPtr->getPage() );
	SlidePersistPtr pMasterPersistPtr( pSlidePersistPtr->getMasterPersist() );
	if ( pMasterPersistPtr.get() )
	{
		const rtl::OUString sLayout( RTL_CONSTASCII_USTRINGPARAM( "Layout" ) );
		uno::Reference< beans::XPropertySet > xSet( xSlide, uno::UNO_QUERY_THROW );
		xSet->setPropertyValue(	sLayout, Any( pMasterPersistPtr->getLayoutFromValueToken() ) );
	}
	while( xSlide->getCount() )
	{
		Reference< drawing::XShape > xShape;
		xSlide->getByIndex(0) >>= xShape;
		xSlide->remove( xShape );
	}

	Reference< XPropertySet > xPropertySet( xSlide, UNO_QUERY );
	if ( xPropertySet.is() )
	{
		static const OUString sWidth( RTL_CONSTASCII_USTRINGPARAM( "Width" ) );
		static const OUString sHeight( RTL_CONSTASCII_USTRINGPARAM( "Height" ) );
		awt::Size& rPageSize( pSlidePersistPtr->isNotesPage() ? maNotesSize : maSlideSize );
		xPropertySet->setPropertyValue( sWidth, Any( rPageSize.Width ) );
		xPropertySet->setPropertyValue( sHeight, Any( rPageSize.Height ) );
	}
	pSlidePersistPtr->setPath( rSlideFragmentPath );
    return mxFilter->importFragment( rxSlideFragmentHandler, rSlideFragmentPath );
}

} }

