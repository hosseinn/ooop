/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textparagraphproperties.cxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 13:35:31 $
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

#include <com/sun/star/text/XNumberingRulesSupplier.hpp>
#include <com/sun/star/container/XIndexReplace.hpp>

#include "oox/core/propertyset.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/drawingml/textparagraphproperties.hxx"
#include "tokens.hxx"



using rtl::OUString;
using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::container;

namespace oox { namespace drawingml {

TextParagraphProperties::TextParagraphProperties()
: maTextCharacterPropertiesPtr( new TextCharacterProperties() )
, mnLevel( 0 )
{
}
TextParagraphProperties::~TextParagraphProperties()
{
}

void TextParagraphProperties::pushToPropSet( const Reference < XPropertySet > & xPropSet ) const
{
    PropertySet aPropSet( xPropSet );
	Sequence< OUString > aNames;
	Sequence< Any > aValues;

//		 maTextParagraphPropertyMap.dump_debug("TextParagraph paragraph props");
	maTextParagraphPropertyMap.makeSequence( aNames, aValues );
	aPropSet.setProperties( aNames, aValues );

    maTextCharacterPropertiesPtr->pushToPropSet(aPropSet.getXPropertySet());

	Reference< XIndexReplace > xNumRule;
	Any aValue;
	const rtl::OUString sNumberingRules( OUString::intern( RTL_CONSTASCII_USTRINGPARAM( "NumberingRules" ) ) );
	aValue = xPropSet->getPropertyValue( sNumberingRules );
	aValue >>= xNumRule;

	OSL_ENSURE( xNumRule.is(), "can't get Numbering rules");
	if( xNumRule.is() )
	{
		Sequence< PropertyValue > aPropSeq;
		OSL_TRACE("OOX: BulletListProps for level %d", getLevel());
//			pProps->getBulletListPropertyMap().dump_debug();
		maBulletListPropertyMap.makeSequence( aPropSeq );
		if( aPropSeq.hasElements() )
		{
			OSL_TRACE("OOX: bullet props inserted at level %d", getLevel());
			xNumRule->replaceByIndex( getLevel(), makeAny( aPropSeq ) );
		}
		xPropSet->setPropertyValue( sNumberingRules, makeAny( xNumRule ) );
	}
}


} }
