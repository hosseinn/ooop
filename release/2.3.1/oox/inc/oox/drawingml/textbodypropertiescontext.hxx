/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textbodypropertiescontext.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: hub $ $Date: 2007/05/16 11:57:25 $
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

#ifndef OOX_DRAWINGML_TEXTBODYPROPERTIESCONTEXT_HXX
#define OOX_DRAWINGML_TEXTBODYPROPERTIESCONTEXT_HXX

#include "oox/drawingml/shape.hxx"

#ifndef OOX_CORE_CONTEXT_HXX
#include "oox/core/context.hxx"
#endif

namespace oox { namespace drawingml {

class TextBodyPropertiesContext : public ::oox::core::Context
{
public:
	TextBodyPropertiesContext( const ::oox::core::ContextRef& xParent, 
														 const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& xAttributes, 
														 oox::drawingml::ShapePtr pShapePtr );

    virtual void SAL_CALL endFastElement( ::sal_Int32 Element ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

protected:
	::oox::drawingml::ShapePtr mpShapePtr;
};

} }

#endif  //  OOX_DRAWINGML_TEXTBODYPROPERTIESCONTEXT_HXX
