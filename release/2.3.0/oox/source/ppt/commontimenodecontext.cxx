/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: commontimenodecontext.cxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: hub $ $Date: 2007/07/12 20:00:56 $
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

#include <algorithm>
#include <boost/cast.hpp>

#include "comphelper/anytostring.hxx"
#include "cppuhelper/exc_hlp.hxx"
#include <osl/diagnose.h>

#include <com/sun/star/animations/XTimeContainer.hpp>
#include <com/sun/star/animations/XAnimationNode.hpp>
#include <com/sun/star/animations/AnimationFill.hpp>
#include <com/sun/star/animations/AnimationRestart.hpp>
#include <com/sun/star/animations/Timing.hpp>
#include <com/sun/star/presentation/TextAnimationType.hpp>
#include <com/sun/star/presentation/EffectPresetClass.hpp>
#include <com/sun/star/presentation/EffectNodeType.hpp>

#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"
#include "oox/core/fragmenthandler.hxx"
#include "oox/ppt/pptimport.hxx"

#include "commontimenodecontext.hxx"
#include "tokens.hxx"

using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::animations;
using namespace ::com::sun::star::presentation;
using namespace ::com::sun::star::xml::sax;


using ::rtl::OUString;
using ::com::sun::star::beans::NamedValue;

namespace oox { namespace ppt {

	
	
	CommonTimeNodeContext::CommonTimeNodeContext( const FragmentHandlerRef& xHandler, 
																								sal_Int32  aElement, 
																								const Reference< XFastAttributeList >& xAttribs, 
																								const TimeNodePtr & pNode )
		: TimeNodeContext( xHandler, aElement, xAttribs, pNode )
			, mbIterate( false )
	{
		OSL_TRACE( "OOX: CommonTimeNodeContext::CommonTimeNodeContext" );
		AttributeList attribs( xAttribs );
		sal_Int32 nInt; // some temporary int value for float conversions
		
		NodePropertyMap & aProps = pNode->getNodeProperties();
		PropertyMap & aUserData = pNode->getUserData();

		nInt = attribs.getInteger( XML_accel, 0 );
		aProps[ NP_ACCELERATION ] = makeAny( ( static_cast<double>(nInt) / 100000.0 ) );
		
		if( attribs.hasAttribute( XML_afterEffect ) ) 
		{
			aUserData[ CREATE_OUSTRING( "after-effect" ) ] 
				= makeAny( attribs.getBool( XML_afterEffect, false ) );
		}
		aProps[ NP_AUTOREVERSE ] = makeAny( attribs.getBool( XML_autoRev, false ) );
		
		// TODO
		if( attribs.hasAttribute( XML_bldLvl ) ) 
		{
			attribs.getInteger( XML_bldLvl, 0 );
		}
		nInt = attribs.getInteger( XML_decel, 0 );
		aProps[ NP_DECELERATE ] <<= ( static_cast<double>(nInt) / 100000.0 );
		// TODO
		if( attribs.hasAttribute( XML_display ) ) 
		{
			attribs.getBool( XML_display, false );
		}
		if( attribs.hasAttribute( XML_dur ) )
		{
			// ST_TLTime
			double fDuration = static_cast<double>(attribs.getUnsignedInteger(  XML_dur, 0 ) ) / 1000.0;
			aProps[ NP_DURATION ] <<= makeAny( fDuration );
		}
		// TODO
		if( attribs.hasAttribute( XML_evtFilter ) )
		{
			xAttribs->getOptionalValue( XML_evtFilter );
		}
		// ST_TLTimeNodeFillType
		if( attribs.hasAttribute( XML_fill ) )
		{
			nInt = xAttribs->getOptionalValueToken( XML_fill, 0 );
			if( nInt != 0 )
			{
				sal_Int16 nEnum;
				switch( nInt )
				{
				case XML_remove:
					nEnum = AnimationFill::REMOVE;
					break;
				case XML_freeze:
					nEnum = AnimationFill::FREEZE;
					break;
				case XML_hold:
					nEnum = AnimationFill::HOLD;
					break;
				case XML_transition:
					nEnum = AnimationFill::TRANSITION;
					break;
				default:
					nEnum = AnimationFill::DEFAULT;
					break;
				}
				aProps[ NP_FILL ] <<=  (sal_Int16)nEnum;
			}
		}
		if( attribs.hasAttribute( XML_grpId ) )
		{
			attribs.getUnsignedInteger( XML_grpId, 0 );
		}
		// ST_TLTimeNodeID
		if( attribs.hasAttribute( XML_id ) ) 
		{
			sal_uInt32 nId = attribs.getUnsignedInteger( XML_id, 0 );
			pNode->setId( nId );
		}
		// ST_TLTimeNodeMasterRelation TODO
		xAttribs->getOptionalValue( XML_masterRel );
		// TODO
		if( attribs.hasAttribute( XML_nodePh ) )
		{
			attribs.getBool( XML_nodePh, false );
		}
		// ST_TLTimeNodeType
		nInt = xAttribs->getOptionalValueToken( XML_nodeType, 0 );
		if( nInt != 0 )
		{
			sal_Int16 nEnum;
			switch( nInt )
			{
			case XML_clickEffect:
			case XML_clickPar:
				nEnum = EffectNodeType::ON_CLICK;
				break;
			case XML_withEffect:
			case XML_withGroup:
				nEnum = EffectNodeType::WITH_PREVIOUS;
				break;
			case XML_mainSeq:
				nEnum = EffectNodeType::MAIN_SEQUENCE;
				break;
			case XML_interactiveSeq:
				nEnum = EffectNodeType::INTERACTIVE_SEQUENCE;
				break;
			case XML_afterGroup:
			case XML_afterEffect:
				nEnum = EffectNodeType::AFTER_PREVIOUS;
				break;
			case XML_tmRoot:
				nEnum = EffectNodeType::TIMING_ROOT;
				break;
			default:
				nEnum = EffectNodeType::DEFAULT;
				break;
			}
			aUserData[ CREATE_OUSTRING( "node-type" ) ] <<= nEnum;
		}

		// ST_TLTimeNodePresetClassType
		nInt = xAttribs->getOptionalValueToken( XML_presetClass, 0 );
		if( nInt != 0 )
		{
			// TODO put that in a function
			sal_Int16 nEnum;
			switch( nInt )
			{
			case XML_entr:
				nEnum = EffectPresetClass::ENTRANCE;
				break;
			case XML_exit:
				nEnum = EffectPresetClass::EXIT;
				break;
			case XML_emph:
				nEnum = EffectPresetClass::EMPHASIS;
				break;
			case XML_path:
				nEnum = EffectPresetClass::MOTIONPATH;
				break;
			case XML_verb:
				// TODO check that the value below is correct
				nEnum = EffectPresetClass::CUSTOM;
				break;
			case XML_mediacall:
				nEnum = EffectPresetClass::MEDIACALL;
				break;
			default:
				nEnum = 0;
				break;
			}
			aUserData[ CREATE_OUSTRING( "preset-class" ) ] = makeAny( nEnum );
		}
		if( attribs.hasAttribute( XML_presetID ) )
		{
			aUserData[ CREATE_OUSTRING(  "preset-id" ) ] 
				=	 makeAny( attribs.getInteger( XML_presetID, 0 ) );
		}
		if( attribs.hasAttribute( XML_presetSubtype ) )
		{
			aUserData[ CREATE_OUSTRING( "preset-sub-type" ) ]
				= makeAny( attribs.getInteger( XML_presetSubtype, 0 ) );
		}
		if( attribs.hasAttribute( XML_repeatCount ) )
		{
			// ST_TLTime
			double fCount = (double)attribs.getUnsignedInteger( XML_repeatCount, 1000 ) / 1000.0;
			/* see pptinanimation */
			aProps[ NP_REPEATCOUNT ] <<= (fCount < ((float)3.40282346638528860e+38)) ? makeAny( (double)fCount ) : makeAny( Timing_INDEFINITE );
		}
		if( attribs.hasAttribute(  XML_repeatDur ) ) 
		{
			// ST_TLTime
			double fDuration = (double)attribs.getUnsignedInteger( XML_repeatDur, 0 ) / 1000.0;
			aProps[ NP_REPEATDURATION ] <<= makeAny( fDuration );
		}
		// ST_TLTimeNodeRestartType
		nInt = xAttribs->getOptionalValueToken( XML_restart, 0 );
		if( nInt != 0 )
		{
			// TODO put that in a function
			sal_Int16 nEnum;
			switch( nInt )
			{
			case XML_always:
				nEnum = AnimationRestart::ALWAYS;
				break;
			case XML_whenNotActive:
				nEnum = AnimationRestart::WHEN_NOT_ACTIVE;
				break;
			case XML_never:
				nEnum = AnimationRestart::NEVER;
				break;
			default:
				nEnum =	AnimationRestart::DEFAULT;
				break;
			}
			aProps[ NP_RESTART ] <<= (sal_Int16)nEnum;
		}
		// ST_Percentage TODO
		xAttribs->getOptionalValue( XML_spd /*"10000" */ );
		// ST_TLTimeNodeSyncType TODO
		xAttribs->getOptionalValue( XML_syncBehavior );
		// TODO (string)
		xAttribs->getOptionalValue( XML_tmFilter );
	}


	CommonTimeNodeContext::~CommonTimeNodeContext( ) throw ( )
	{
	}


	void SAL_CALL CommonTimeNodeContext::endFastElement( sal_Int32 aElement ) throw ( SAXException, RuntimeException)
	{
		if( aElement == mnElement )
		{
			// TODO I don't know what to do with the other types...
			if( maEndSyncValue.mnCondition == ( NMSP_PPT|XML_rtn ) )
			{
				mpNode->getNodeProperties()[ NP_ENDSYNC ] <<= maEndSyncValue.maValue;
			}
			// TODO push maStCondList and maEndCondList
		}
		else if( aElement == ( NMSP_PPT|XML_iterate ) )
		{
			mbIterate = false;
		}
	}


	Reference< XFastContextHandler > SAL_CALL CommonTimeNodeContext::createFastChildContext( ::sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw ( SAXException, RuntimeException )
	{
		Reference< XFastContextHandler > xRet;

		switch ( aElementToken )
		{
		case NMSP_PPT|XML_childTnLst:
		case NMSP_PPT|XML_subTnLst:
			xRet.set( new TimeNodeListContext( getHandler(), mpNode->getChilds() ) );
			break;

		case NMSP_PPT|XML_stCondLst:
			xRet.set( new CondListContext( getHandler(), aElementToken, xAttribs, mpNode, maStCondList ) );
			break;
		case NMSP_PPT|XML_endCondLst:
			xRet.set( new CondListContext( getHandler(), aElementToken, xAttribs, mpNode, maEndCondList ) );
			break;

		case NMSP_PPT|XML_endSync:
			xRet.set( new CondContext( getHandler(), xAttribs, mpNode, maEndSyncValue ) );
			break;
		case NMSP_PPT|XML_iterate:
		{
			sal_Int32 nVal = xAttribs->getOptionalValueToken( XML_type, XML_el );
			if( nVal != 0 )
			{
				// TODO put that in a function
				sal_Int16 nEnum;
				switch( nVal )
				{
				case XML_el:
					nEnum =	TextAnimationType::BY_PARAGRAPH;
					break;
				case XML_lt:
					nEnum = TextAnimationType::BY_LETTER;
					break;
				case XML_wd:
					nEnum = TextAnimationType::BY_WORD;
					break;
				default:
					// default is BY_WORD. See Ppt97Animation::GetTextAnimationType()
					// in sd/source/filter/ppt/ppt97animations.cxx:297
					nEnum = TextAnimationType::BY_WORD;
					break;
				}
				mpNode->getNodeProperties()[ NP_ITERATETYPE ] <<= nEnum;
			}
			// in case of exception we ignore the whole tag.
			AttributeList attribs( xAttribs );		
			// TODO what to do with this
			/*bool bBackwards =*/ attribs.getBool( XML_backwards, false );
			mbIterate = true;
			break;
		}
		case NMSP_PPT|XML_tmAbs:
			if( mbIterate )
			{
				AttributeList attribs( xAttribs );			
				double fTime = attribs.getUnsignedInteger( XML_val, 0 );
				// time in ms. property is in % TODO
				mpNode->getNodeProperties()[ NP_ITERATEINTERVAL ] <<= fTime;
			}
			break;
		case NMSP_PPT|XML_tmPct:
			if( mbIterate )
			{
				AttributeList attribs( xAttribs );			
				double fPercent = (double)attribs.getUnsignedInteger( XML_val, 0 ) / 100000.0;
				mpNode->getNodeProperties()[ NP_ITERATEINTERVAL ] <<= fPercent;
			}
			break;
		default:
			break;
		}

		if( !xRet.is() )
			xRet.set( this );

		return xRet;
	}

} }
