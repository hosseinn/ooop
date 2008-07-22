/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textparagraphpropertiescontext.hxx,v $
 *
 *  $Revision: 1.1.2.8 $
 *
 *  last change: $Author: hub $ $Date: 2007/06/15 23:30:16 $
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

#ifndef OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIESCONTEXT_HXX
#define OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIESCONTEXT_HXX

#include <list>

#include <com/sun/star/style/TabStop.hpp>

#include "oox/drawingml/textparagraphproperties.hxx"
#include "oox/drawingml/textfont.hxx"
#include "oox/drawingml/textspacing.hxx"

#ifndef OOX_CORE_CONTEXT_HXX
#include "oox/core/context.hxx"
#endif

namespace oox { namespace drawingml {


class BulletListProps
{
public:
	BulletListProps( );
	bool is() const;
	void pushToProperties( 	::oox::core::PropertyMap& rProps );
	void setBulletChar( const ::rtl::OUString & sChar );
	void setStartAt( sal_Int32 nStartAt )
        { mnStartAt = static_cast< sal_Int16 >( nStartAt ); }
	void setType( sal_Int32 nType );
	void setNone( );
	void setSuffixParenBoth();
	void setSuffixParenRight();
	void setSuffixPeriod();
	void setSuffixNone();
	void setSuffixMinusRight();
	void setBulletSize(sal_Int16 nSize);
	void setFontSize(sal_Int16 nSize);

	sal_Int32	                                   mnBulletColor;
	bool                                         mbHasBulletColor;
	bool                                         mbBulletColorFollowText;
	bool                                         mbBulletFontFollowText;
	TextFont                                     maBulletFont;
	::rtl::OUString                              msBulletChar;
	sal_Int16                                    mnStartAt;
	sal_Int16                                    mnNumberingType;
	::rtl::OUString                              msNumberingPrefix;
	::rtl::OUString                              msNumberingSuffix;
	sal_Int16                                    mnSize;
	sal_Int16                                    mnFontSize;
};



class TextParagraphPropertiesContext : public ::oox::core::Context
{
public:
	TextParagraphPropertiesContext( const ::oox::core::ContextRef& xParent,
		const com::sun::star::uno::Reference< com::sun::star::xml::sax::XFastAttributeList >& rXAttributes,
			oox::drawingml::TextParagraphPropertiesPtr pTextParagraphPropertiesPtr );
	~TextParagraphPropertiesContext();

	virtual void SAL_CALL endFastElement( ::sal_Int32 Element ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler > SAL_CALL createFastChildContext( ::sal_Int32 Element, const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& Attribs ) throw (::com::sun::star::xml::sax::SAXException, ::com::sun::star::uno::RuntimeException);

protected:
	::oox::drawingml::TextParagraphPropertiesPtr mpTextParagraphPropertiesPtr;
	TextSpacing                                  maLineSpacing;
	TextSpacing                                  maSpaceBefore;
	TextSpacing                                  maSpaceAfter;
	BulletListProps                              maBulletListProps;
	::std::list< ::com::sun::star::style::TabStop >  maTabList;
};

} }

#endif  //  OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIESCONTEXT_HXX