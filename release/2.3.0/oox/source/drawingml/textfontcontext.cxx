/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textfontcontext.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/05/22 07:53:26 $
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


#include "oox/core/attributelist.hxx"
#include "oox/drawingml/textfont.hxx"
#include "textfontcontext.hxx"

using namespace ::oox::core;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::uno;


namespace oox { namespace drawingml {

	TextFontContext::TextFontContext( const FragmentHandlerRef& xHandler, sal_Int32 nToken,
																		const Reference< XFastAttributeList >& xAttributes,
																		TextFont & aFont )
		: Context( xHandler )
			, mnToken( nToken )
			, maFont( aFont )
	{
		AttributeList attribs( xAttributes );

		maFont.msTypeface = xAttributes->getValue( XML_typeface );
		maFont.msPanose = xAttributes->getOptionalValue( XML_panose );
		maFont.mnPitch = attribs.getInteger( XML_pitchFamily, 0 );
		maFont.mnCharset = attribs.getInteger( XML_charset, 1 );
	}

	void TextFontContext::endFastElement( sal_Int32 nElement )
		throw ( SAXException, RuntimeException )
	{
		if( nElement == mnToken )
		{

		}
	}

    Reference< XFastContextHandler > TextFontContext::createFastChildContext( ::sal_Int32 /*aElement*/, const Reference< XFastAttributeList >& /*Attribs*/ )
		throw ( SAXException, RuntimeException )
	{
		Reference< XFastContextHandler > xRet;
//       switch( aElement )
//       {
//
//       }
		if ( !xRet.is() )
			xRet.set( this );
		return xRet;
	}


} }
