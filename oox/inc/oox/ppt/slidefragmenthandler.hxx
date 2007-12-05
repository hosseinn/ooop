/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: slidefragmenthandler.hxx,v $
 *
 *  $Revision: 1.1.2.11 $
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

#ifndef OOX_PPT_SLIDEFRAGMENTHANDLER
#define OOX_PPT_SLIDEFRAGMENTHANDLER

#include <com/sun/star/drawing/XDrawPage.hpp>
#include "oox/vml/drawingfragmenthandler.hxx"
#include "oox/core/fragmenthandler.hxx"
#include "oox/core/propertymap.hxx"
#include "oox/ppt/slidepersist.hxx"

#include <stack>
#include <vector>
#include <map>

namespace oox { namespace ppt {

class SlideFragmentHandler : public ::oox::core::FragmentHandler
{
public:
	SlideFragmentHandler( const oox::core::XmlFilterRef& xFilter, const ::rtl::OUString& rFragmentPath, oox::ppt::SlidePersistPtr pPersistPtr, const oox::ppt::ShapeLocation eShapeLocation ) throw();
	virtual ~SlideFragmentHandler() throw();

    virtual void SAL_CALL endDocument(  ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

protected:
	::oox::ppt::SlidePersistPtr		mpSlidePersistPtr;
	::oox::ppt::ShapeLocation		meShapeLocation;

private:
	::oox::core::RelationsRef		mxRelations;

	::rtl::OUString					maSlideName;
	::oox::core::PropertyMap		maSlideProperties;
};

} }

#endif // OOX_PPT_SLIDEFRAGMENTHANDLER