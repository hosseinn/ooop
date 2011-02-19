/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: timeanimvaluecontext.hxx,v $
 *
 *  $Revision: 1.1.2.1 $
 *
 *  last change: $Author: hub $ $Date: 2007/05/02 17:04:59 $
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



#ifndef OOX_PPT_TIMEANIMVALUELISTCONTEXT
#define OOX_PPT_TIMEANIMVALUELISTCONTEXT

#include "oox/core/context.hxx"
#include "oox/ppt/animationspersist.hxx"

namespace oox { namespace ppt {

	/** CT_TLTimeAnimateValueList */
	class TimeAnimValueListContext 
		: public ::oox::core::Context
	{
	public:
		TimeAnimValueListContext( const ::oox::core::FragmentHandlerRef& xHandler,
															const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttribs, 
															TimeAnimationValueList & aTavList );

		~TimeAnimValueListContext( );

		virtual void SAL_CALL endFastElement( sal_Int32 aElement ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
		
		virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 aElementToken, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& /*xAttribs*/ ) throw ( ::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException );


	private:
		TimeAnimationValueList & maTavList;
		bool                     mbInValue;
	};




} }

#endif