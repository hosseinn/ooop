/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textparagraphproperties.hxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: hub $ $Date: 2007/06/19 19:50:45 $
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

#ifndef OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIES_HXX
#define OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIES_HXX

#include <com/sun/star/beans/XPropertySet.hpp>
#include "oox/drawingml/textcharacterproperties.hxx"

namespace oox { namespace drawingml {

class TextParagraphProperties;

typedef boost::shared_ptr< TextParagraphProperties > TextParagraphPropertiesPtr;

class TextParagraphProperties
{
public:

	TextParagraphProperties();
    ~TextParagraphProperties();

	void                                setLevel( sal_Int16 nLevel ) { mnLevel = nLevel; }
	sal_Int16                           getLevel( ) const 
		{ return mnLevel; }
	::oox::core::PropertyMap&						getTextParagraphPropertyMap()
		{ return maTextParagraphPropertyMap; }
	::oox::core::PropertyMap&           getBulletListPropertyMap()
		{ return maBulletListPropertyMap; }
	::oox::drawingml::TextCharacterPropertiesPtr	getTextCharacterProperties()
		{ return maTextCharacterPropertiesPtr; }
	void                                pushToPropSet( const ::com::sun::star::uno::Reference < ::com::sun::star::beans::XPropertySet > & xPropSet ) const;
protected:

	TextCharacterPropertiesPtr		maTextCharacterPropertiesPtr;
	::oox::core::PropertyMap		  maTextParagraphPropertyMap;
	::oox::core::PropertyMap	    maBulletListPropertyMap;
	sal_Int16                     mnLevel;
};

} }

#endif  //  OOX_DRAWINGML_TEXTPARAGRAPHPROPERTIES_HXX
