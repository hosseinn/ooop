/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: shape.cxx,v $
 *
 *  $Revision: 1.1.2.22 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/30 11:19:19 $
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

#include "oox/drawingml/shape.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

#ifndef _SOLAR_H
#include <tools/solar.h>
#endif
#include <com/sun/star/graphic/XGraphic.hpp>
#include <com/sun/star/container/XNamed.hpp>
#include <com/sun/star/beans/XMultiPropertySet.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/drawing/HomogenMatrix3.hpp>
#include <com/sun/star/text/XText.hpp>
#include <basegfx/point/b2dpoint.hxx>
#include <basegfx/matrix/b2dhommatrix.hxx>

using rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star;
using namespace ::com::sun::star::awt;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::frame;
using namespace ::com::sun::star::text;
using namespace ::com::sun::star::drawing;

namespace oox { namespace drawingml {

Shape::Shape( const sal_Char* pServiceName )
: mpTextListStyle( new TextListStyle() )
, mnSubType( 0 )
, mnIndex( 0 )
, mnRotation( 0 )
, mbFlipH( false )
, mbFlipV( false )
{
	if ( pServiceName )
		msServiceName = OUString::createFromAscii( pServiceName );
	setDefaults();
}
Shape::~Shape()
{
}

void Shape::setDefaults()
{
	const OUString sTextAutoGrowHeight( RTL_CONSTASCII_USTRINGPARAM( "TextAutoGrowHeight" ) );
	const OUString sTextWordWrap( RTL_CONSTASCII_USTRINGPARAM( "TextWordWrap" ) );
	const OUString sTextLeftDistance( RTL_CONSTASCII_USTRINGPARAM( "TextLeftDistance" ) );
	const OUString sTextUpperDistance( RTL_CONSTASCII_USTRINGPARAM( "TextUpperDistance" ) );
	const OUString sTextRightDistance( RTL_CONSTASCII_USTRINGPARAM( "TextRightDistance" ) );
	const OUString sTextLowerDistance( RTL_CONSTASCII_USTRINGPARAM( "TextLowerDistance" ) );
	maShapeProperties[ sTextAutoGrowHeight ] <<= sal_False;
	maShapeProperties[ sTextWordWrap ] <<= sal_True;
	maShapeProperties[ sTextLeftDistance ]  <<= static_cast< sal_Int32 >( 250 );
	maShapeProperties[ sTextUpperDistance ] <<= static_cast< sal_Int32 >( 125 );
	maShapeProperties[ sTextRightDistance ] <<= static_cast< sal_Int32 >( 250 );
	maShapeProperties[ sTextLowerDistance ] <<= static_cast< sal_Int32 >( 125 );
}

void Shape::addShape( const Reference< XModel > &rxModel, const oox::drawingml::ThemePtr pThemePtr,
				std::map< ::rtl::OUString, ShapePtr > & aShapeMap, const Reference< XShapes >& rxShapes, const awt::Rectangle* pShapeRect )
{
    try
    {
		rtl::OUString sServiceName( msServiceName );
		if( sServiceName.getLength() )
		{
			Reference< XShape > xShape( createAndInsert( sServiceName, rxModel, pThemePtr, rxShapes, pShapeRect ) );

			if( msId.getLength() )
			{
				aShapeMap[ msId ] = shared_from_this();
			}

			// if this is a group shape, we have to add also each child shape
			Reference< XShapes > xShapes( xShape, UNO_QUERY );
			if ( xShapes.is() )
				addChilds( *this, rxModel, pThemePtr, aShapeMap, xShapes, pShapeRect ? *pShapeRect : awt::Rectangle( maPosition.X, maPosition.Y, maSize.Width, maSize.Height ) );
		}
	}
	catch( const Exception&  )
	{
	}
}

void Shape::applyShapeReference( const oox::drawingml::Shape& rReferencedShape )
{
	mpTextBody = TextBodyPtr( new TextBody( *rReferencedShape.mpTextBody.get() ) );
	maShapeProperties = rReferencedShape.maShapeProperties;
	maCustomShapeGeometry = rReferencedShape.maCustomShapeGeometry;
	mpTextListStyle = TextListStylePtr( new TextListStyle( *rReferencedShape.mpTextListStyle.get() ) );
	maShapeStylesColorMap = rReferencedShape.maShapeStylesColorMap;
	maShapeStylesIndexMap = rReferencedShape.maShapeStylesIndexMap;
	maSize = rReferencedShape.maSize;
	maPosition = rReferencedShape.maPosition;
	mnRotation = rReferencedShape.mnRotation;
	mbFlipH = rReferencedShape.mbFlipH;
	mbFlipV = rReferencedShape.mbFlipV;
}

// for group shapes, the following method is also adding each child
void Shape::addChilds( Shape& rMaster, const Reference< XModel > &rxModel, const oox::drawingml::ThemePtr pThemePtr,
					  std::map< ::rtl::OUString, ShapePtr > & aShapeMap, const Reference< XShapes >& rxShapes, const awt::Rectangle& rClientRect )
{
	// first the global child union needs to be calculated
    sal_Int32 nGlobalLeft  = SAL_MAX_INT32;
    sal_Int32 nGlobalRight = SAL_MIN_INT32;
    sal_Int32 nGlobalTop   = SAL_MAX_INT32;
    sal_Int32 nGlobalBottom= SAL_MIN_INT32;
	std::vector< ShapePtr >::iterator aIter( rMaster.maChilds.begin() );
	while( aIter != rMaster.maChilds.end() )
	{
		sal_Int32 l = (*aIter)->maPosition.X;
		sal_Int32 t = (*aIter)->maPosition.Y;
		sal_Int32 r = l + (*aIter)->maSize.Width;
		sal_Int32 b = t + (*aIter)->maSize.Height;
		if ( nGlobalLeft > l )
			nGlobalLeft = l;
		if ( nGlobalRight < r )
			nGlobalRight = r;
		if ( nGlobalTop > t )
			nGlobalTop = t;
		if ( nGlobalBottom < b )
			nGlobalBottom = b;
		aIter++;
	}
	aIter = rMaster.maChilds.begin();
	while( aIter != rMaster.maChilds.end() )
	{
        Rectangle aShapeRect;
        Rectangle* pShapeRect = 0;
        if ( ( nGlobalLeft != SAL_MAX_INT32 ) && ( nGlobalRight != SAL_MIN_INT32 ) && ( nGlobalTop != SAL_MAX_INT32 ) && ( nGlobalBottom != SAL_MIN_INT32 ) )
		{
			sal_Int32 nGlobalWidth = nGlobalRight - nGlobalLeft;
			sal_Int32 nGlobalHeight = nGlobalBottom - nGlobalTop;
			if ( nGlobalWidth && nGlobalHeight )
			{
				double fWidth = (*aIter)->maSize.Width;
				double fHeight= (*aIter)->maSize.Height;
				double fXScale = (double)rClientRect.Width / (double)nGlobalWidth;
				double fYScale = (double)rClientRect.Height / (double)nGlobalHeight;
				aShapeRect.X = static_cast< sal_Int32 >( ( ( (*aIter)->maPosition.X - nGlobalLeft ) * fXScale ) + rClientRect.X );
				aShapeRect.Y = static_cast< sal_Int32 >( ( ( (*aIter)->maPosition.Y - nGlobalTop  ) * fYScale ) + rClientRect.Y );
				fWidth *= fXScale;
				fHeight *= fYScale;
				aShapeRect.Width = static_cast< sal_Int32 >( fWidth );
				aShapeRect.Height = static_cast< sal_Int32 >( fHeight );
				pShapeRect = &aShapeRect;
			}
		}
		(*aIter++)->addShape( rxModel, pThemePtr, aShapeMap, rxShapes, pShapeRect );
	}
}

Reference< XShape > Shape::createAndInsert( const rtl::OUString& rServiceName, const ::com::sun::star::uno::Reference< ::com::sun::star::frame::XModel > &rxModel,
								const oox::drawingml::ThemePtr pThemePtr, const ::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShapes >& rxShapes,
								const awt::Rectangle* pShapeRect )
{
	basegfx::B2DHomMatrix aTransformation;

	awt::Size aSize( pShapeRect ? awt::Size( pShapeRect->Width, pShapeRect->Height ) : maSize );
	awt::Point aPosition( pShapeRect ? awt::Point( pShapeRect->X, pShapeRect->Y ) : maPosition );
	if( aSize.Width != 1 || aSize.Height != 1)
	{
		// take care there are no zeros used by error
		aTransformation.scale(
			aSize.Width ? aSize.Width : 1.0,
			aSize.Height ? aSize.Height : 1.0);
	}

	if( mbFlipH || mbFlipV || mnRotation != 0)
	{
		// calculate object's center
		basegfx::B2DPoint aCenter(0.5, 0.5);
		aCenter *= aTransformation;

		// center object at origin
		aTransformation.translate( -aCenter.getX(), -aCenter.getY() );

		if( mbFlipH || mbFlipV)
		{
			// mirror around object's center
			aTransformation.scale( mbFlipH ? -1.0 : 1.0, mbFlipV ? -1.0 : 1.0 );
		}

		if( mnRotation != 0 )
		{
			// rotate around object's center
			aTransformation.rotate( -F_PI180 * ( (double)mnRotation / 60000.0 ) );
		}

		// move object back from center
		aTransformation.translate( aCenter.getX(), aCenter.getY() );
	}

	if( aPosition.X != 0 || aPosition.Y != 0)
	{
		// if global position is used, add it to transformation
		aTransformation.translate( aPosition.X, aPosition.Y );
	}

	// now set transformation for this object
	HomogenMatrix3 aMatrix;

	aMatrix.Line1.Column1 = aTransformation.get(0,0);
	aMatrix.Line1.Column2 = aTransformation.get(0,1);
	aMatrix.Line1.Column3 = aTransformation.get(0,2);

	aMatrix.Line2.Column1 = aTransformation.get(1,0);
	aMatrix.Line2.Column2 = aTransformation.get(1,1);
	aMatrix.Line2.Column3 = aTransformation.get(1,2);

	aMatrix.Line3.Column1 = aTransformation.get(2,0);
	aMatrix.Line3.Column2 = aTransformation.get(2,1);
	aMatrix.Line3.Column3 = aTransformation.get(2,2);

	static const OUString sTransformation(RTL_CONSTASCII_USTRINGPARAM("Transformation"));
	maShapeProperties[ sTransformation ] <<= aMatrix;

	Reference< lang::XMultiServiceFactory > xServiceFact( rxModel, UNO_QUERY_THROW );
	Reference< drawing::XShape > xShape( xServiceFact->createInstance( rServiceName ), UNO_QUERY_THROW );
	if( xShape.is() )
	{
		if ( rServiceName == OUString::createFromAscii( "com.sun.star.drawing.GraphicObjectShape" ) )
		{
			Reference< graphic::XGraphic > xGraphic;
			static const OUString sFillBitmap( RTL_CONSTASCII_USTRINGPARAM( "FillBitmap" ) );
			if ( maShapeProperties[ sFillBitmap ] >>= xGraphic )
			{
				maShapeProperties.erase( sFillBitmap );
				static const OUString sGraphic( RTL_CONSTASCII_USTRINGPARAM( "Graphic" ) );
				maShapeProperties[ sGraphic ] <<= xGraphic;
			}
		}
		if( msName.getLength() )
		{
			Reference< container::XNamed > xNamed( xShape, UNO_QUERY );
			if( xNamed.is() )
				xNamed->setName( msName );
		}
		rxShapes->add( xShape );
		mxShape = xShape;

		setShapeStyles( pThemePtr );

		// applying properties
		if( !maShapeProperties.empty() )
		{
/* sj: setPropertyValues is throwing "unknown property value" exceptions, therefore I already wrote an issue

			Reference< XMultiPropertySet > xMSet( xShape, UNO_QUERY );
			if( xMSet.is() )
			{
				Sequence< OUString > aNames;
				Sequence< Any > aValues;
				maShapeProperties.makeSequence( aNames, aValues );
				xMSet->setPropertyValues( aNames,  aValues);
			}
			else
*/
			{
				uno::Reference< beans::XPropertySet > xSet( xShape, uno::UNO_QUERY_THROW );
				uno::Reference< beans::XPropertySetInfo > xInfo( xSet->getPropertySetInfo() );

				for( PropertyMap::const_iterator aIter( maShapeProperties.begin() ); aIter != maShapeProperties.end(); aIter++ )
				{
					if ( xInfo->hasPropertyByName( (*aIter).first ) )
						xSet->setPropertyValue( (*aIter).first, (*aIter).second );
				}
			}



		}
		if( !maCustomShapeGeometry.empty() && ( rServiceName == OUString::createFromAscii( "com.sun.star.drawing.CustomShape" ) ) )
		{
			Reference< XPropertySet > xSet( xShape, UNO_QUERY );
			if( xSet.is() )
			{
				// converting the vector to a sequence
				Sequence< PropertyValue > aSeq;
				maCustomShapeGeometry.makeSequence( aSeq );
				static const rtl::OUString sCustomShapeGeometry( RTL_CONSTASCII_USTRINGPARAM( "CustomShapeGeometry" ) );
				xSet->setPropertyValue( sCustomShapeGeometry, Any( aSeq ) );
			}
		}
		
		// in some cases, we don't have any text body.
		if( getTextBody() )
		{
			Reference < XText > xText( xShape, UNO_QUERY );
			if ( xText.is() )	// not every shape is supporting an XText interface (e.g. GroupShape)
			{
				Reference < XTextCursor > xAt = xText->createTextCursor();
				getTextBody()->insertAt( xText, xAt, rxModel );
			}
		}
	}
	return xShape;
}

// the properties of rDest which are not part of rSource are being put into rSource
void addMissingProperties( const oox::core::PropertyMap& rSource, oox::core::PropertyMap& rDest )
{
	oox::core::PropertyMap::const_iterator aSourceIter( rSource.begin() );
	while( aSourceIter != rSource.end() )
	{
		if ( rDest.find( (*aSourceIter ).first ) == rDest.end() )
			rDest[ (*aSourceIter).first ] <<= (*aSourceIter).second; 
		aSourceIter++;
	}
}

// merging styles, if a shape property is not set, we have to set the shape style property
void Shape::setShapeStyles( const oox::drawingml::ThemePtr pThemePtr )
{
	std::map< ShapeStyle, sal_Int32 >::const_iterator aShapeStylesColorIter( getShapeStylesColor().begin() );
	std::map< ShapeStyle, rtl::OUString >::const_iterator aShapeStylesIndexIter( getShapeStylesIndex().begin() );
	while( aShapeStylesColorIter != getShapeStylesColor().end() )
	{
		switch( (*aShapeStylesColorIter).first )
		{
			case oox::drawingml::SHAPESTYLE_ln :
			{
				const rtl::OUString sLineColor( OUString::intern( RTL_CONSTASCII_USTRINGPARAM( "LineColor" ) ) );
				if ( maShapeProperties.find( sLineColor ) == maShapeProperties.end() )
					maShapeProperties[ sLineColor ] <<= (*aShapeStylesColorIter).second; 
			}
			break;
			case oox::drawingml::SHAPESTYLE_fill :
			{
				const rtl::OUString sFillColor( OUString::intern( RTL_CONSTASCII_USTRINGPARAM( "FillColor" ) ) );
				if ( maShapeProperties.find( sFillColor ) == maShapeProperties.end() )
					maShapeProperties[ sFillColor ] <<= (*aShapeStylesColorIter).second; 
			}
			case oox::drawingml::SHAPESTYLE_effect :
			break;
			case oox::drawingml::SHAPESTYLE_font :
			{
				const rtl::OUString sCharColor( OUString::intern( RTL_CONSTASCII_USTRINGPARAM( "CharColor" ) ) );
				if ( maShapeProperties.find( sCharColor ) == maShapeProperties.end() )
					maShapeProperties[ sCharColor ] <<= (*aShapeStylesColorIter).second;
			}
			break;
		}
		aShapeStylesColorIter++;
	}
	while( aShapeStylesIndexIter != getShapeStylesIndex().end() )
	{
		const rtl::OUString sIndex( (*aShapeStylesIndexIter).second );
		sal_uInt32 nIndex( sIndex.toInt32() );
		if ( nIndex-- )
		{
			switch( (*aShapeStylesIndexIter).first )
			{
				case oox::drawingml::SHAPESTYLE_ln :
				{
					const std::vector< oox::core::PropertyMap >& rLineStyleList( pThemePtr->getLineStyleList() );
					if ( rLineStyleList.size() > nIndex )
						addMissingProperties( rLineStyleList[ nIndex ], maShapeProperties );
				}
				break;
				case oox::drawingml::SHAPESTYLE_fill :
				{
					const std::vector< oox::core::PropertyMap >& rFillStyleList( pThemePtr->getFillStyleList() );
					if ( rFillStyleList.size() > nIndex )
						addMissingProperties( rFillStyleList[ nIndex ], maShapeProperties );
				}
				break;
				case oox::drawingml::SHAPESTYLE_effect :
				case oox::drawingml::SHAPESTYLE_font :
				break;
			}
		}
		aShapeStylesIndexIter++;
	}
}

void Shape::setTextBody(const TextBodyPtr & pTextBody)
{
	mpTextBody = pTextBody;
}


TextBodyPtr Shape::getTextBody()
{
	return mpTextBody;
}


} }
