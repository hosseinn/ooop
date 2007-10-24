/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: slidefragmenthandler.cxx,v $
 *
 *  $Revision: 1.1.2.25 $
 *
 *  last change: $Author: sj $ $Date: 2007/09/04 17:06:17 $
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

#include <com/sun/star/beans/XMultiPropertySet.hpp>
#include <com/sun/star/container/XNamed.hpp>

#include <oox/ppt/backgroundproperties.hxx>
#include "oox/ppt/slidefragmenthandler.hxx"
#include "oox/ppt/slidetimingcontext.hxx"
#include "oox/ppt/slidetransitioncontext.hxx"
#include "oox/ppt/slidemastertextstylescontext.hxx"
#include "oox/ppt/pptshapegroupcontext.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"
#include "oox/ppt/pptshape.hxx"


using rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::container;

#define C2U(x) OUString( RTL_CONSTASCII_USTRINGPARAM( x ) )

namespace oox { namespace ppt {

SlideFragmentHandler::SlideFragmentHandler( const oox::core::XmlFilterRef& xFilter, const ::rtl::OUString& rFragmentPath, oox::ppt::SlidePersistPtr pPersistPtr, const oox::ppt::ShapeLocation eShapeLocation ) throw()
: FragmentHandler( xFilter, rFragmentPath )
, mpSlidePersistPtr( pPersistPtr )
, meShapeLocation( eShapeLocation )
{
	RelationPtr pVMLDrawing( getRelations()->getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" ) ) );
	if ( pVMLDrawing )
	{
		OUString aVMLDrawingFragmentPath( resolveRelativePath( rFragmentPath, pVMLDrawing->msTarget ) );
		if ( aVMLDrawingFragmentPath.getLength() )
		{
			Reference< XFastDocumentHandler > xVMLDrawingFragmentHandler(
				new oox::vml::DrawingFragmentHandler( mxFilter, aVMLDrawingFragmentPath, pPersistPtr->getDrawing() ) );
            mxFilter->importFragment( xVMLDrawingFragmentHandler, aVMLDrawingFragmentPath );
		}
	}
}

SlideFragmentHandler::~SlideFragmentHandler() throw()
{
}

Reference< XFastContextHandler > SlideFragmentHandler::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
	case NMSP_PPT|XML_sldMaster:		// CT_SlideMaster
	case NMSP_PPT|XML_handoutMaster:	// CT_HandoutMaster
	case NMSP_PPT|XML_sld:				// CT_CommonSlideData
	case NMSP_PPT|XML_notes:			// CT_NotesSlide
	case NMSP_PPT|XML_notesMaster:		// CT_NotesMaster
		break;
	case NMSP_PPT|XML_cSld:				// CT_CommonSlideData
		maSlideName = xAttribs->getOptionalValue(XML_name);
		break;

	case NMSP_PPT|XML_spTree:			// CT_GroupShape
        xRet.set( new PPTShapeGroupContext( mpSlidePersistPtr, meShapeLocation, this, aElementToken, mpSlidePersistPtr->getShapes(),
			oox::drawingml::ShapePtr( new PPTShape( meShapeLocation, "com.sun.star.drawing.GroupShape" ) ) ) );
		break;

	case NMSP_PPT|XML_timing: // CT_SlideTiming
		xRet.set( new SlideTimingContext( this, mpSlidePersistPtr->getTimeNodeList() ) );
		break;
	case NMSP_PPT|XML_transition: // CT_SlideTransition
		xRet.set( new SlideTransitionContext( this, xAttribs, maSlideProperties ) );
		break;

	// BackgroundGroup
	case NMSP_PPT|XML_bgPr:				// CT_BackgroundProperties
		xRet.set( new BackgroundPropertiesContext( this, maSlideProperties ) );
		break;
	case NMSP_PPT|XML_bgRef:			// a:CT_StyleMatrixReference
		break;

	case NMSP_PPT|XML_clrMap:			// CT_ColorMapping
	case NMSP_PPT|XML_clrMapOvr:		// CT_ColorMappingOverride
	case NMSP_PPT|XML_sldLayoutIdLst:	// CT_SlideLayoutIdList
		break;
	case NMSP_PPT|XML_txStyles:			// CT_SlideMasterTextStyles
		xRet.set( new SlideMasterTextStylesContext( this, mpSlidePersistPtr ) );
		break;
	case NMSP_PPT|XML_custDataLst:		// CT_CustomerDataList
	case NMSP_PPT|XML_tagLst:			// CT_TagList
		break;
	}

	if( !xRet.is() )
		xRet.set(this);

	return xRet;
}

void SAL_CALL SlideFragmentHandler::endDocument(  ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException)
{
	try
	{
		Reference< XDrawPage > xSlide( mpSlidePersistPtr->getPage() );
		if( !maSlideProperties.empty() )
		{
/* sj: setPropertyValues is throwing "unknown property value" exceptions, therefore I already wrote an issue

			uno::Reference< beans::XMultiPropertySet > xMSet( xSlide, uno::UNO_QUERY );
			if( xMSet.is() )
			{
				uno::Sequence< OUString > aNames;
				uno::Sequence< uno::Any > aValues;
				maSlideProperties.makeSequence( aNames, aValues );
				xMSet->setPropertyValues( aNames,  aValues);
			}
			else
*/
			{
				uno::Reference< beans::XPropertySet > xSet( xSlide, uno::UNO_QUERY_THROW );
				uno::Reference< beans::XPropertySetInfo > xInfo( xSet->getPropertySetInfo() );

				for( PropertyMap::const_iterator aIter( maSlideProperties.begin() ); aIter != maSlideProperties.end(); aIter++ )
				{
					if ( xInfo->hasPropertyByName( (*aIter).first ) )
						xSet->setPropertyValue( (*aIter).first, (*aIter).second );
				}
			}
		}
		if ( maSlideName.getLength() )
		{
			Reference< XNamed > xNamed( xSlide, UNO_QUERY );
			if( xNamed.is() )
				xNamed->setName( maSlideName );
		}
	}
	catch( uno::Exception& )
	{
        OSL_ENSURE( false,
			(rtl::OString("oox::ppt::SlideFragmentHandler::EndElement(), "
					"exception caught: ") +
			rtl::OUStringToOString(
				comphelper::anyToString( cppu::getCaughtException() ),
				RTL_TEXTENCODING_UTF8 )).getStr() );
	}
}

} }

