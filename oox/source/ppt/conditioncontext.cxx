/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: conditioncontext.cxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: hub $ $Date: 2007/07/07 04:22:27 $
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
#include <osl/diagnose.h>


#include <com/sun/star/animations/XTimeContainer.hpp>
#include <com/sun/star/animations/XAnimationNode.hpp>
#include <com/sun/star/animations/AnimationEndSync.hpp>

#include "oox/core/namespaces.hxx"
#include "oox/core/fragmenthandler.hxx"
#include "oox/core/attributelist.hxx"
#include "oox/core/context.hxx"
#include "oox/ppt/animationspersist.hxx"

#include "timetargetelementcontext.hxx"
#include "tokens.hxx"

using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::animations;

#include "conditioncontext.hxx"

namespace oox { namespace ppt {

	CondContext::CondContext( const FragmentHandlerRef & xHandler, const Reference< XFastAttributeList >& xAttribs,
														const TimeNodePtr & pNode, AnimationCondition & aValue )
		:  TimeNodeContext( xHandler, NMSP_PPT|XML_cond, xAttribs, pNode )
			 , maCond( aValue )
	{
		AttributeList attribs( xAttribs );
		// TODO figure out if it is 0 or -1 when the attribute is missing
		maCond.mnDelay = attribs.getInteger( XML_delay, 0 );
		maCond.mnEvent = xAttribs->getOptionalValueToken( XML_evt, 0 );				
	}

	CondContext::~CondContext( ) throw( )
	{
	}
	
	void SAL_CALL CondContext::endFastElement( sal_Int32 aElement ) throw ( SAXException, RuntimeException)
	{
		if( aElement == ( NMSP_PPT|XML_cond ) ) 
		{
			// TODO 
		}
	}

	Reference< XFastContextHandler > SAL_CALL CondContext::createFastChildContext( ::sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw ( SAXException, RuntimeException )
	{
		Reference< XFastContextHandler > xRet;
		
		switch( aElementToken )
		{
		case NMSP_PPT|XML_rtn:
		{
			// ST_TLTriggerRuntimeNode { first, last, all }
			sal_Int32 aTok;
			sal_Int16 nEnum;
			aTok = xAttribs->getOptionalValueToken( XML_val, XML_first );
			switch( aTok )
			{
			case XML_first:
				nEnum = AnimationEndSync::FIRST;
				break;
			case XML_last:
				nEnum = AnimationEndSync::LAST;
				break;
			case XML_all:
				nEnum = AnimationEndSync::ALL;
				break;
			default: 
				break;
			}
			maCond.mnCondition = aElementToken;
			maCond.maValue = makeAny( nEnum );
			break;
		}
		case NMSP_PPT|XML_tn:
		{
			maCond.mnCondition = aElementToken;
			AttributeList attribs( xAttribs );
			sal_uInt32 nId = attribs.getUnsignedInteger( XML_val, 0 );
			maCond.maValue = makeAny( nId );
			break;	
		}
		case NMSP_PPT|XML_tgtEl:
			// CT_TLTimeTargetElement
			maCond.mnCondition = aElementToken;
			xRet.set( new TimeTargetElementContext( getHandler(), maCond.getTarget() ) );
			break;
		default:
			break;
		}
		
		if( !xRet.is() )
			xRet.set( this );
		
		return xRet;
		
	}
	




	/** CT_TLTimeConditionList */
	CondListContext::CondListContext( const FragmentHandlerRef & xHandler, sal_Int32  aElement, 
																		const Reference< XFastAttributeList >& xAttribs, 
																		const TimeNodePtr & pNode,
																		AnimationConditionList & aCondList )
		: TimeNodeContext( xHandler, aElement, xAttribs, pNode )
			, maConditions( aCondList )
	{
	}


	CondListContext::~CondListContext( ) throw( )
	{
	}


	void SAL_CALL CondListContext::endFastElement( sal_Int32 aElement ) throw ( SAXException, RuntimeException)
	{
		if( aElement == mnElement )
		{
			// TODO flush the con
			OSL_TRACE( "OOX: number of conditions. %d", maConditions.size() );
		}
	}

		
	Reference< XFastContextHandler > CondListContext::createFastChildContext( ::sal_Int32 aElement, const Reference< XFastAttributeList >& xAttribs ) throw ( SAXException, RuntimeException )
	{
		Reference< XFastContextHandler > xRet;

		switch( aElement )
		{
		case NMSP_PPT|XML_cond:
			// add a condition to the list
			maConditions.push_back( AnimationCondition() );
			xRet.set( new CondContext( getHandler(), xAttribs, mpNode, maConditions.back() ) );
			break;
		default:
			break;
		}

		if( !xRet.is() )
			xRet.set( this );
		
		return xRet;
	}


} }

