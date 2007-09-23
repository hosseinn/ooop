/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textcharacterproperties.hxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: hub $ $Date: 2007/08/23 14:37:50 $
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

#ifndef OOX_DRAWINGML_TEXTCHARACTERPROPERTIES_HXX
#define OOX_DRAWINGML_TEXTCHARACTERPROPERTIES_HXX

#include "oox/core/propertymap.hxx"
#include <com/sun/star/beans/XPropertySet.hpp>
#include <boost/shared_ptr.hpp>

namespace oox { namespace drawingml {

class TextCharacterProperties;

typedef boost::shared_ptr< TextCharacterProperties > TextCharacterPropertiesPtr;

class TextCharacterProperties
{
public:

	TextCharacterProperties();
	~TextCharacterProperties();

	::oox::core::PropertyMap&		getTextCharacterPropertyMap() { return maTextCharacterPropertyMap; }
	::oox::core::PropertyMap&       getHyperlinkPropertyMap()     { return maHyperlinkPropertyMap; }
	void pushToPropSet( const ::com::sun::star::uno::Reference < ::com::sun::star::beans::XPropertySet > & xPropSet ) const;
	void pushToUrlFieldPropSet( const ::com::sun::star::uno::Reference < ::com::sun::star::beans::XPropertySet > & xPropSet ) const;
protected:
	::oox::core::PropertyMap		maTextCharacterPropertyMap;
	::oox::core::PropertyMap		maHyperlinkPropertyMap;
};

} }

#endif  //  OOX_DRAWINGML_TEXTCHARACTERPROPERTIES_HXX
