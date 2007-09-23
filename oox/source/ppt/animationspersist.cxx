/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: animationspersist.cxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: dr $ $Date: 2007/07/09 14:34:01 $
 *
 *  The Contents of this file are made available subject to
 *  the terms of GNU Lesser General Public License Version 2.1.
 *
 *
 *    GNU Lesser General Public License Version 2.1
 *    =============================================
 *    Copyright 2007 by Sun Microsystems, Inc.
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



#include "oox/ppt/animationspersist.hxx"

#include <rtl/ustring.hxx>
#include <com/sun/star/uno/Any.hxx>
#include <com/sun/star/drawing/XShape.hpp>
#include <com/sun/star/text/XText.hpp>
#include <com/sun/star/presentation/ParagraphTarget.hpp>
#include <com/sun/star/presentation/ShapeAnimationSubType.hpp>

#include "oox/drawingml/shape.hxx"

#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::presentation;
using namespace ::com::sun::star::drawing;
using namespace ::com::sun::star::text;

namespace oox { namespace ppt {

	void ShapeTargetElement::convert( ::com::sun::star::uno::Any & rTarget, sal_Int16 & rSubType ) const
	{
		switch(mnType)
		{
		case XML_subSp:
			rSubType = ShapeAnimationSubType::AS_WHOLE;
			break;
		case XML_bg:
			rSubType = ShapeAnimationSubType::ONLY_BACKGROUND;
			break;
		case XML_txEl:
		{
			ParagraphTarget aParaTarget;
			Reference< XShape > xShape;
			rTarget >>= xShape;
			aParaTarget.Shape = xShape;
			rSubType = ShapeAnimationSubType::ONLY_TEXT;

			Reference< XText > xText( xShape, UNO_QUERY );
			if( xText.is() )
			{
				switch(mnRangeType)
				{
				case XML_charRg:
					// TODO calculate the corresponding paragraph for the text range....

					break;
				case XML_pRg:
                    aParaTarget.Paragraph = static_cast< sal_Int16 >( maRange.start );
					// TODO what to do with more than one.
					break;
				}
				rTarget = makeAny( aParaTarget );
			}
			break;
		}
		default:
			break;
		}
	}


	void AnimTargetElement::convert(NodePropertyMap & aProperties, const SlidePersistPtr & pSlide) const
	{
		// see sd/source/files/ppt/pptinanimations.cxx:3191 (in importTargetElementContainer())
		switch(mnType)
		{
		case XML_inkTgt:
			// TODO
			break;
		case XML_sldTgt:
			// TODO
			break;
		case XML_sndTgt:
			aProperties[ NP_TARGET ] = makeAny(msValue);
			break;
		case XML_spTgt:
		{
			sal_Int16 nSubType;
			Any rTarget;
			::oox::drawingml::ShapePtr pShape = pSlide->getShape(msValue);
			OSL_ENSURE( pShape, "failed to locate Shape");
			if( pShape )
			{
				Reference< XShape > xShape( pShape->getXShape() );
				OSL_ENSURE( xShape.is(), "fail to get XShape from shape" );
				if( xShape.is() )
				{
					rTarget <<= xShape;
					maShapeTarget.convert(rTarget, nSubType);
					aProperties[ NP_SUBITEM ] = makeAny( nSubType );
					aProperties[ NP_TARGET ] = rTarget;
				}
			}
			break;
		}
		default:
			break;
		}
	}


} }


