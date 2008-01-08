/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textcharacterproperties.cxx,v $
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


#include "oox/core/propertyset.hxx"
#include "oox/drawingml/textcharacterproperties.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

#include <com/sun/star/beans/XPropertySet.hpp>

using rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;

namespace oox { namespace drawingml {

TextCharacterProperties::TextCharacterProperties()
{
}
TextCharacterProperties::~TextCharacterProperties()
{
}

void TextCharacterProperties::pushToPropSet( const Reference < XPropertySet > & xPropSet ) const
{
    PropertySet aPropSet( xPropSet );
	Sequence< OUString > aNames;
	Sequence< Any > aValues;

//		 maTextCharacterPropertyMap.dump_debug("TextCharacter props");
	maTextCharacterPropertyMap.makeSequence( aNames, aValues );
	aPropSet.setProperties( aNames, aValues );
}



void TextCharacterProperties::pushToUrlFieldPropSet( const Reference < XPropertySet > & xPropSet ) const
{
    PropertySet aPropSet( xPropSet );
	Sequence< OUString > aNames;
	Sequence< Any > aValues;

	maHyperlinkPropertyMap.makeSequence( aNames, aValues );
	aPropSet.setProperties( aNames, aValues );
}



} }
