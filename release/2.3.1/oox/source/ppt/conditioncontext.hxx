/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: conditioncontext.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: hub $ $Date: 2007/07/05 15:18:40 $
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


#ifndef OOX_PPT_CONDITIONCONTEXT
#define OOX_PPT_CONDITIONCONTEXT

#include "oox/core/fragmenthandler.hxx"
#include "oox/ppt/timenode.hxx"
#include "oox/ppt/timenodelistcontext.hxx"
#include "oox/ppt/animationspersist.hxx"

namespace oox { namespace ppt {


	/** CT_TLTimeCondition */
	class CondContext
		: public TimeNodeContext
	{
	public:
		CondContext( const ::oox::core::FragmentHandlerRef & xHandler, 
								 const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttribs, 
								 const TimeNodePtr & pNode, AnimationCondition & aValue );
		~CondContext( ) throw( );
		virtual void SAL_CALL endFastElement( sal_Int32 /*aElement*/ ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
		virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 aElementToken, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttribs ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException );
		
	private:
		AnimationCondition & maCond;
	};



	/** CT_TLTimeConditionList */
	class CondListContext
		: public TimeNodeContext
	{
	public:
		CondListContext( const ::oox::core::FragmentHandlerRef & xHandler, 
										 sal_Int32  aElement, 
										 const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttribs, 
										 const TimeNodePtr & pNode, AnimationConditionList & aCondList );
		~CondListContext( ) throw( );

		virtual void SAL_CALL endFastElement( sal_Int32 aElement ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

		virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 aElementToken, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& /*xAttribs*/ ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException );

	private:
		AnimationConditionList & maConditions;
	};


} }


#endif