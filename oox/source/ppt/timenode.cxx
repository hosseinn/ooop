/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: timenode.cxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 13:35:31 $
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


#include <boost/bind.hpp>

#include <comphelper/processfactory.hxx>

#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/animations/XAnimateColor.hpp>
#include <com/sun/star/animations/XAnimateMotion.hpp>
#include <com/sun/star/animations/XAnimateTransform.hpp>
#include <com/sun/star/animations/XCommand.hpp>
#include <com/sun/star/animations/XIterateContainer.hpp>
#include <com/sun/star/animations/XAnimationNodeSupplier.hpp>
#include <com/sun/star/animations/XTimeContainer.hpp>
#include <com/sun/star/animations/AnimationNodeType.hpp>

#include "oox/core/helper.hxx"
#include "oox/ppt/timenode.hxx"


using ::rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::animations;
using namespace ::com::sun::star::frame;

namespace oox { namespace ppt {

		OUString TimeNode::getServiceName( sal_Int16 nNodeType )
		{
			OUString sServiceName;
			switch( nNodeType )
			{
			case AnimationNodeType::PAR:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.IterateContainer");
//				sServiceName = CREATE_OUSTRING("com.sun.star.animations.ParallelTimeContainer");
				break;
			case AnimationNodeType::SEQ:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.SequenceTimeContainer");
				break;
			case AnimationNodeType::ANIMATE:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.Animate");
				break;
			case AnimationNodeType::ANIMATECOLOR:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.AnimateColor");
				break;
			case AnimationNodeType::TRANSITIONFILTER:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.TransitionFilter");
				break;
			case AnimationNodeType::ANIMATEMOTION:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.AnimateMotion");
				break;
			case AnimationNodeType::ANIMATETRANSFORM:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.AnimateTransform");
				break;
			case AnimationNodeType::COMMAND:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.Command");
				break;
			case AnimationNodeType::SET:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.AnimateSet");
				break;
			case AnimationNodeType::AUDIO:
				sServiceName = CREATE_OUSTRING("com.sun.star.animations.Audio");
				break;
			default:
				OSL_TRACE( "OOX: uhandled type %x", nNodeType );
				break;
			}
			return sServiceName;
		}



	TimeNode::TimeNode( sal_Int16 nNodeType )
		: mnNodeType( nNodeType )
	{
	}


	TimeNode::~TimeNode()
	{
	}

	void TimeNode::addNode( const Reference< XModel > &rxModel,
													const Reference< XAnimationNode >& rxNode, const SlidePersistPtr & pSlide )
	{
		try {
			OUString sServiceName = getServiceName( mnNodeType );
			Reference< XAnimationNode > xNode = createAndInsert(sServiceName, rxModel, rxNode );

			if( xNode.is() )
			{
				if( msId.getLength() )
				{
					pSlide->getAnimNodesMap()[ msId ] = xNode;
					OSL_TRACE( "OOX: added node to map" );
				}

				if( mpTarget )
				{
					mpTarget->convert( maNodeProperties, pSlide );
				}

				Sequence< NamedValue > aUserDataSeq;
				maUserData.makeSequence(aUserDataSeq);
				if( aUserDataSeq.getLength() )
				{
					maNodeProperties[ NP_USERDATA ] = makeAny(aUserDataSeq);
				}

				Reference< XAnimate > xAnimate( xNode, UNO_QUERY );
				Reference< XAnimateColor > xAnimateColor( xNode, UNO_QUERY );
				Reference< XAnimateMotion > xAnimateMotion( xNode, UNO_QUERY );
				Reference< XAnimateTransform > xAnimateTransform( xNode, UNO_QUERY );
				Reference< XCommand > xCommand( xNode, UNO_QUERY );
				Reference< XIterateContainer > xIterateContainer( xNode, UNO_QUERY );
				sal_Int16 nInt16;
				sal_Bool bBool;
				double fDouble;
				OUString sString;
				Sequence< NamedValue > aSeq;

				for( NodePropertyMap::const_iterator aIter( maNodeProperties.begin() );
						 aIter != maNodeProperties.end(); aIter++ )
				{
					switch( aIter->first )
					{
					case NP_TO:
						if( xAnimate.is() )
							xAnimate->setTo( aIter->second );
						break;
					case NP_FROM:
						if( xAnimate.is() )
							xAnimate->setFrom( aIter->second );
						break;
					case NP_BY:
						if( xAnimate.is() )
							xAnimate->setBy( aIter->second );
						break;
					case NP_TARGET:
						if( xAnimate.is() )
							xAnimate->setTarget( aIter->second );
						break;
					case NP_SUBITEM:
						if( xAnimate.is() )
						{
							aIter->second >>= nInt16;
							xAnimate->setSubItem( nInt16 );
						}
						break;
					case NP_ATTRIBUTENAME:
						if( xAnimate.is() )
						{
							aIter->second >>= sString;
							xAnimate->setAttributeName( sString );
						}
					case NP_CALCMODE:
						if( xAnimate.is() )
						{
							aIter->second >>= nInt16;
							xAnimate->setCalcMode( nInt16 );
						}
						break;
 					case NP_KEYTIMES:
						if( xAnimate.is() )
						{
							Sequence<double> aKeyTimes;
							aIter->second >>= aKeyTimes;
							xAnimate->setKeyTimes(aKeyTimes);
						}
						break;
					case NP_VALUES:
						if( xAnimate.is() )
						{
							Sequence<Any> aValues;
							aIter->second >>= aValues;
							xAnimate->setValues(aValues);
						}
						break;
					case NP_FORMULA:
						if( xAnimate.is() )
						{
							aIter->second >>= sString;
							xAnimate->setFormula(sString);
						}
						break;
					case NP_COLORINTERPOLATION:
						if( xAnimateColor.is() )
						{
							aIter->second >>= nInt16;
							xAnimateColor->setColorInterpolation( nInt16 );
						}
						break;
					case NP_DIRECTION:
						if( xAnimateColor.is() )
						{
							aIter->second >>= bBool;
							xAnimateColor->setDirection( bBool );
						}
						break;
					case NP_PATH:
						if( xAnimateMotion.is() )
							xAnimateMotion->setPath( aIter->second );
						break;
					case NP_TRANSFORMTYPE:
						if( xAnimateTransform.is() )
						{
							aIter->second >>= nInt16;
							xAnimateTransform->setTransformType( nInt16 );
						}
						break;
					case NP_USERDATA:
						aIter->second >>= aSeq;
						xNode->setUserData( aSeq );
						break;
					case NP_ACCELERATION:
						aIter->second >>= fDouble;
						xNode->setAcceleration( fDouble );
						break;
					case NP_DECELERATE:
						aIter->second >>= fDouble;
						xNode->setDecelerate( fDouble );
						break;
					case NP_AUTOREVERSE:
						aIter->second >>= bBool;
						xNode->setAutoReverse( bBool );
						break;
					case NP_DURATION:
						xNode->setDuration( aIter->second );
						break;
					case NP_FILL:
						aIter->second >>= nInt16;
						xNode->setFill( nInt16 );
						break;
					case NP_REPEATCOUNT:
						xNode->setRepeatCount( aIter->second );
						break;
					case NP_REPEATDURATION:
						xNode->setRepeatDuration( aIter->second );
						break;
					case NP_RESTART:
						aIter->second >>= nInt16;
						xNode->setRestart( nInt16 );
						break;
					case NP_ENDSYNC:
						xNode->setEndSync(aIter->second);
						break;
					case NP_COMMAND:
						if( xCommand.is() )
						{
							aIter->second >>= nInt16;
							xCommand->setCommand( nInt16 );
						}
						break;
					case NP_PARAMETER:
						if( xCommand.is() )
							xCommand->setParameter( aIter->second );
						break;
					case NP_ITERATETYPE:
						if( xIterateContainer.is() )
						{
							aIter->second >>= nInt16;
							xIterateContainer->setIterateType( nInt16 );
						}
						break;
					case NP_ITERATEINTERVAL:
						if( xIterateContainer.is() )
						{
							aIter->second >>= fDouble;
							xIterateContainer->setIterateInterval( fDouble );
						}
						break;
					default:
						break;
					}
				}

				if( mnNodeType == AnimationNodeType::TRANSITIONFILTER )
				{

					Reference< XTransitionFilter > xFilter( xNode, UNO_QUERY );
					maTransitionFilter.setTransitionFilterProperties( xFilter );
				}

				std::for_each( maChilds.begin(), maChilds.end(),
											 boost::bind(&TimeNode::addNode, _1, rxModel, boost::ref(xNode),
																	 boost::ref(pSlide) ) );
			}
			else
			{
				OSL_TRACE( "OOX: TimeNode::addNode() coudln't create xNode" );
			}
		}
		catch( const Exception& e )
		{
			OSL_TRACE( "OOX: exception raised in TimeNode::addNode() - %s",
								 OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr() );
		}
	}


    Reference< XAnimationNode > TimeNode::createAndInsert( const OUString& rServiceName, const Reference< XModel > &/*rxModel*/,
																							 const Reference< XAnimationNode >& rxNode )
	{
		try {
			Reference< XAnimationNode > xNode ( ::comphelper::getProcessServiceFactory()->createInstance(rServiceName ),
																					UNO_QUERY_THROW );
			Reference< XTimeContainer > xParentContainer( rxNode, UNO_QUERY_THROW );

			xParentContainer->appendChild( xNode );
			return xNode;
		}
		catch( const Exception& e )
		{
			OSL_TRACE( "OOX: exception raised in TimeNode::createAndInsert() trying to create a service %s = %s",
								 OUStringToOString( rServiceName, RTL_TEXTENCODING_ASCII_US ).getStr(),
								 OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr() );
		}

		return Reference< XAnimationNode >();
	}


	void 	TimeNode::setId( sal_Int32 nId )
	{
		msId = OUString::valueOf(nId);
	}

	void TimeNode::setTo( const Any & aTo )
	{
		maNodeProperties[ NP_TO ] = aTo;
	}


	void TimeNode::setFrom( const Any & aFrom )
	{
		maNodeProperties[ NP_FROM ] = aFrom;
	}

	void TimeNode::setBy( const Any & aBy )
	{
		maNodeProperties[ NP_BY ] = aBy;
	}


} }
