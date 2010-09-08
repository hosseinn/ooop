/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textcharacterpropertiescontext.hxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: hub $ $Date: 2007/06/13 17:59:14 $
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

#ifndef OOX_DRAWINGML_TEXTCHARACTERPROPERTIESCONTEXT_HXX
#define OOX_DRAWINGML_TEXTCHARACTERPROPERTIESCONTEXT_HXX

#include "oox/drawingml/textparagraphproperties.hxx"
#include "oox/drawingml/textfont.hxx"

#ifndef OOX_CORE_CONTEXT_HXX
#include "oox/core/context.hxx"
#endif
#include "oox/core/propertymap.hxx"

namespace oox { namespace drawingml {

class TextCharacterPropertiesContext : public ::oox::core::Context
{
public:
	TextCharacterPropertiesContext( const ::oox::core::ContextRef& xParent,
		const com::sun::star::uno::Reference< com::sun::star::xml::sax::XFastAttributeList >& rXAttributes,
			oox::drawingml::TextCharacterPropertiesPtr pTextCharacterPropertiesPtr );
	~TextCharacterPropertiesContext();

	virtual void SAL_CALL endFastElement( ::sal_Int32 Element ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

protected:
	::oox::drawingml::TextCharacterPropertiesPtr mpTextCharacterPropertiesPtr;
	bool      mbHasHighlightColor;
	sal_Int32	mnHighlightColor;
	sal_Int32	mnCharColor;
	TextFont  maLatinFont;
	TextFont  maAsianFont;
	TextFont  maComplexFont;
	TextFont  maSymbolFont;
	bool mbHasUnderline;
	bool mbUnderlineLineFollowText;
	bool mbHasUnderlineColor;
	bool mbUnderlineFillFollowText;
	sal_Int32	mnUnderlineColor;
};

} }

#endif  //  OOX_DRAWINGML_TEXTCHARACTERPROPERTIESCONTEXT_HXX