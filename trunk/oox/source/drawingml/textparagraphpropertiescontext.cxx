/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: textparagraphpropertiescontext.cxx,v $
 *
 *  $Revision: 1.1.2.18 $
 *
 *  last change: $Author: hub $ $Date: 2007/08/30 18:07:06 $
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

#include <com/sun/star/style/NumberingType.hpp>
#include <com/sun/star/text/WritingMode.hpp>
#include <com/sun/star/text/HoriOrientation.hpp>
#include <com/sun/star/awt/FontDescriptor.hpp>

#include "oox/drawingml/colorchoicecontext.hxx"
#include "oox/drawingml/textcharacterpropertiescontext.hxx"
#include "oox/drawingml/textparagraphpropertiescontext.hxx"
#include "oox/drawingml/drawingmltypes.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"
#include "textfontcontext.hxx"
#include "textspacingcontext.hxx"
#include "texttabstoplistcontext.hxx"
#include "tokens.hxx"

using ::rtl::OUString;
using namespace ::oox::core;
using ::com::sun::star::awt::FontDescriptor;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::style;
using namespace ::com::sun::star::text;

namespace oox { namespace drawingml {


BulletListProps::BulletListProps( )
	: mnBulletColor( 0 )
	, mbHasBulletColor( false )
    , mbBulletColorFollowText( false )
	, mbBulletFontFollowText( false )
	, mnStartAt( 0 )
	, mnNumberingType( NumberingType::NUMBER_NONE )
	, mnSize( 100 ) // default is 100%
	, mnFontSize( 0 )
{
}

bool BulletListProps::is() const
{
	return mnNumberingType != NumberingType::NUMBER_NONE;
}

void BulletListProps::setBulletChar( const ::rtl::OUString & sChar )
{
	mnNumberingType = NumberingType::CHAR_SPECIAL;
	msBulletChar = sChar;
}

void BulletListProps::setNone( )
{
	mnNumberingType = NumberingType::NUMBER_NONE;
	mbHasBulletColor = false;
}

void BulletListProps::setSuffixParenBoth()
{
	msNumberingSuffix = CREATE_OUSTRING( ")" );
	msNumberingPrefix = CREATE_OUSTRING( "(" );
}

void BulletListProps::setSuffixParenRight()
{
	msNumberingSuffix = CREATE_OUSTRING( ")" );
	msNumberingPrefix = OUString();
}

void BulletListProps::setSuffixPeriod()
{
	msNumberingSuffix = CREATE_OUSTRING( "." );
	msNumberingPrefix = OUString();
}

void BulletListProps::setSuffixNone()
{
	msNumberingSuffix = OUString();
	msNumberingPrefix = OUString();
}

void BulletListProps::setSuffixMinusRight()
{
	msNumberingSuffix = CREATE_OUSTRING( "-" );
	msNumberingPrefix = OUString();
}

void BulletListProps::setType( sal_Int32 nType )
{
	OSL_TRACE( "OOX: set list numbering type %d", nType);
	switch( nType )
	{
	case XML_alphaLcParenBoth:
		mnNumberingType = NumberingType::CHARS_LOWER_LETTER;
		setSuffixParenBoth();
		break;
	case XML_alphaLcParenR:
		mnNumberingType = NumberingType::CHARS_LOWER_LETTER;
		setSuffixParenRight();
		break;
	case XML_alphaLcPeriod:
		mnNumberingType = NumberingType::CHARS_LOWER_LETTER;
		setSuffixPeriod();
		break;
	case XML_alphaUcParenBoth:
		mnNumberingType = NumberingType::CHARS_UPPER_LETTER;
		setSuffixParenBoth();
		break;
	case XML_alphaUcParenR:
		mnNumberingType = NumberingType::CHARS_UPPER_LETTER;
		setSuffixParenRight();
		break;
	case XML_alphaUcPeriod:
		mnNumberingType = NumberingType::CHARS_UPPER_LETTER;
		setSuffixPeriod();
		break;
	case XML_arabic1Minus:
	case XML_arabic2Minus:
	case XML_arabicDbPeriod:
	case XML_arabicDbPlain:
		// TODO
		break;
	case XML_arabicParenBoth:
 		mnNumberingType = NumberingType::ARABIC;
		setSuffixParenBoth();
		break;
	case XML_arabicParenR:
 		mnNumberingType = NumberingType::ARABIC;
		setSuffixParenRight();
		break;
	case XML_arabicPeriod:
 		mnNumberingType = NumberingType::ARABIC;
		setSuffixPeriod();
		break;
	case XML_arabicPlain:
 		mnNumberingType = NumberingType::ARABIC;
		setSuffixNone();
		break;
	case XML_circleNumDbPlain:
	case XML_circleNumWdBlackPlain:
	case XML_circleNumWdWhitePlain:
		mnNumberingType = NumberingType::CIRCLE_NUMBER;
		break;
	case XML_ea1ChsPeriod:
		mnNumberingType = NumberingType::NUMBER_UPPER_ZH;
 		setSuffixPeriod();
		break;
	case XML_ea1ChsPlain:
		mnNumberingType = NumberingType::NUMBER_UPPER_ZH;
		setSuffixNone();
		break;
	case XML_ea1ChtPeriod:
		mnNumberingType = NumberingType::NUMBER_UPPER_ZH_TW;
 		setSuffixPeriod();
		break;
	case XML_ea1ChtPlain:
		mnNumberingType = NumberingType::NUMBER_UPPER_ZH_TW;
		setSuffixNone();
		break;
	case XML_ea1JpnChsDbPeriod:
	case XML_ea1JpnKorPeriod:
	case XML_ea1JpnKorPlain:
		break;
	case XML_hebrew2Minus:
		mnNumberingType = NumberingType::CHARS_HEBREW;
		setSuffixMinusRight();
		break;
	case XML_hindiAlpha1Period:
	case XML_hindiAlphaPeriod:
	case XML_hindiNumParenR:
	case XML_hindiNumPeriod:
		// TODO
		break;
	case XML_romanLcParenBoth:
		mnNumberingType = NumberingType::ROMAN_LOWER;
		setSuffixParenBoth();
		break;
	case XML_romanLcParenR:
		mnNumberingType = NumberingType::ROMAN_LOWER;
		setSuffixParenRight();
		break;
	case XML_romanLcPeriod:
		mnNumberingType = NumberingType::ROMAN_LOWER;
 		setSuffixPeriod();
		break;
	case XML_romanUcParenBoth:
		mnNumberingType = NumberingType::ROMAN_UPPER;
		setSuffixParenBoth();
		break;
	case XML_romanUcParenR:
		mnNumberingType = NumberingType::ROMAN_UPPER;
		setSuffixParenRight();
		break;
	case XML_romanUcPeriod:
		mnNumberingType = NumberingType::ROMAN_UPPER;
 		setSuffixPeriod();
		break;
	case XML_thaiAlphaParenBoth:
	case XML_thaiNumParenBoth:
		mnNumberingType = NumberingType::CHARS_THAI;
		setSuffixParenBoth();
		break;
	case XML_thaiAlphaParenR:
	case XML_thaiNumParenR:
		mnNumberingType = NumberingType::CHARS_THAI;
		setSuffixParenRight();
		break;
	case XML_thaiAlphaPeriod:
	case XML_thaiNumPeriod:
		mnNumberingType = NumberingType::CHARS_THAI;
 		setSuffixPeriod();
		break;
	}
}

void BulletListProps::setBulletSize(sal_Int16 nSize)
{
	mnSize = nSize;
}


void BulletListProps::setFontSize(sal_Int16 nSize)
{
	mnFontSize = nSize;
}


void BulletListProps::pushToProperties(	PropertyMap& rProps )
{
	if( msNumberingPrefix.getLength() )
	{
		OSL_TRACE( "OOX: numb prefix found");
		const rtl::OUString sPrefix( CREATE_OUSTRING( "Prefix" ) );
		rProps[ sPrefix ] <<= msNumberingPrefix;
	}
	if( msNumberingSuffix.getLength() )
	{
		OSL_TRACE( "OOX: numb suffix found");
		const rtl::OUString sSuffix( CREATE_OUSTRING( "Suffix" ) );
		rProps[ sSuffix ] <<= msNumberingSuffix;
	}
	if( mnStartAt != 0 )
	{
		const rtl::OUString sStartWith( CREATE_OUSTRING( "StartWith" ) );
		rProps[ sStartWith ] <<= mnStartAt;
	}
	const rtl::OUString sAdjust( CREATE_OUSTRING( "Adjust" ) );
	rProps[ sAdjust ] <<= HoriOrientation::LEFT;
	
	const rtl::OUString sNumberingType( CREATE_OUSTRING( "NumberingType" ) );
	rProps[ sNumberingType ] <<= mnNumberingType;
	FontDescriptor aFontDesc;
	if( mnFontSize != 0 )
	{
		aFontDesc.Height = mnFontSize;
	}
	if( maBulletFont.is() )
	{
		// TODO move the to the TextFont struct.
		sal_Int16 nPitch, nFamily;
		aFontDesc.Name = maBulletFont.msTypeface;
		GetFontPitch( maBulletFont.mnPitch, nPitch, nFamily);
		aFontDesc.Pitch = nPitch;
		aFontDesc.Family = nFamily;
	}
	const rtl::OUString sBulletFont( CREATE_OUSTRING( "BulletFont" ) );
	rProps[ sBulletFont ] <<= aFontDesc;
	if( mnNumberingType == NumberingType::CHAR_SPECIAL )
	{
		if( maBulletFont.is() )
		{
			const rtl::OUString sBulletFontName( CREATE_OUSTRING( "BulletFontName" ) );
			rProps[ sBulletFontName ] <<= maBulletFont.msTypeface;
		}
		const rtl::OUString sBulletChar( CREATE_OUSTRING( "BulletChar" ) );
		rProps[ sBulletChar ] <<= msBulletChar;

		const rtl::OUString sBulletColor( CREATE_OUSTRING( "BulletColor" ) );
		rProps[ sBulletColor ] <<= mnBulletColor;
		

		if( mnSize )
		{
			const rtl::OUString sBulletRelSize( CREATE_OUSTRING( "BulletRelSize" ) );
			rProps[ sBulletRelSize ] <<= (sal_Int16)mnSize;
		}
	}
}



// --------------------------------------------------------------------

// CT_TextParagraphProperties
TextParagraphPropertiesContext::TextParagraphPropertiesContext( const ::oox::core::ContextRef& xParent,
															    const Reference< XFastAttributeList >& xAttribs,
															    oox::drawingml::TextParagraphPropertiesPtr pTextParagraphPropertiesPtr )
	: Context( xParent->getHandler() )
	, mpTextParagraphPropertiesPtr( pTextParagraphPropertiesPtr )
{
	OUString sValue;
	AttributeList attribs( xAttribs );
	OSL_ASSERT( pTextParagraphPropertiesPtr );

	::oox::core::PropertyMap& rPropertyMap( mpTextParagraphPropertiesPtr->getTextParagraphPropertyMap() );

	// ST_TextAlignType
	sal_Int32 nAlign = xAttribs->getOptionalValueToken( XML_algn, XML_l );
	const OUString sParaAdjust( CREATE_OUSTRING( "ParaAdjust" ) );
	rPropertyMap[ sParaAdjust ] <<= GetParaAdjust( nAlign );
	OSL_TRACE( "OOX: para adjust %d", GetParaAdjust( nAlign ));
	// TODO see to do the same with RubyAdjust

	// ST_Coordinate32
	sValue = xAttribs->getOptionalValue( XML_defTabSz );
//	sal_Int32 nDefTabSize = ( sValue.getLength() == 0 ? 0 : GetCoordinate(  sValue ) );
	// TODO

//	bool bEaLineBrk = attribs.getBool( XML_eaLnBrk, true );
	bool bLatinLineBrk = attribs.getBool( XML_latinLnBrk, true );
	const OUString sParaIsHyphenation( CREATE_OUSTRING( "ParaIsHyphenation" ) );
	rPropertyMap[ sParaIsHyphenation ] <<= bLatinLineBrk;
	// TODO see what to do with Asian hyphenation

	// ST_TextFontAlignType
	// TODO
//	sal_Int32 nFontAlign = xAttribs->getOptionalValueToken( XML_fontAlgn, XML_base );

	bool bHangingPunct = attribs.getBool( XML_hangingPunct, false );
	const OUString sParaIsHangingPunctuation( CREATE_OUSTRING( "ParaIsHangingPunctuation" ) );
	rPropertyMap[ sParaIsHangingPunctuation ] <<= bHangingPunct;

  // ST_Coordinate
	sValue = xAttribs->getOptionalValue( XML_indent );
	const OUString sParaFirstLineIndent( CREATE_OUSTRING( "ParaFirstLineIndent" ) );
	rPropertyMap[ sParaFirstLineIndent ] <<= ( sValue.getLength() == 0 ? 0 : GetCoordinate(  sValue ) );

  // ST_TextIndentLevelType
	// -1 is an invalid value and denote the lack of level
	sal_Int32 nLevel = attribs.getInteger( XML_lvl, 0 );
	if( nLevel > 8 || nLevel < 0 )
	{
		nLevel = 0;
	}

	mpTextParagraphPropertiesPtr->setLevel( static_cast< sal_Int16 >( nLevel ) );

	PropertyMap& rBulletListPropertyMap( mpTextParagraphPropertiesPtr->getBulletListPropertyMap() );
	char name[] = "Outline X";
	name[8] = static_cast<char>( '1' + nLevel );
	const OUString sStyleNameValue( rtl::OUString::createFromAscii( name ) );
	const OUString sCharStyleName( CREATE_OUSTRING( "CharStyleName" ) );
	rBulletListPropertyMap[ sCharStyleName ] <<= sStyleNameValue;

	// ST_TextMargin
	// ParaLeftMargin
	sValue = xAttribs->getOptionalValue( XML_marL );
	sal_Int32 nMarL = ( sValue.getLength() == 0 ? 347663 / 360 : GetCoordinate(  sValue ) );
	const OUString sParaLeftMargin( CREATE_OUSTRING( "ParaLeftMargin" ) );
	rPropertyMap[ sParaLeftMargin ] <<= nMarL;

	// ParaRightMargin
	sValue = xAttribs->getOptionalValue( XML_marR );
	sal_Int32 nMarR  = ( sValue.getLength() == 0 ? 0 : GetCoordinate( sValue ) );
	const OUString sParaRightMargin( CREATE_OUSTRING( "ParaRightMargin" ) );
	rPropertyMap[ sParaRightMargin ] <<= nMarR;

	bool bRtl = attribs.getBool( XML_rtl, false );
	const OUString sTextWritingMode( CREATE_OUSTRING( "TextWritingMode" ) );
	rPropertyMap[ sTextWritingMode ] <<= ( bRtl ? WritingMode_RL_TB : WritingMode_LR_TB );
}



TextParagraphPropertiesContext::~TextParagraphPropertiesContext()
{
	PropertyMap& rPropertyMap( mpTextParagraphPropertiesPtr->getTextParagraphPropertyMap() );
	PropertyMap& rBulletListPropertyMap( mpTextParagraphPropertiesPtr->getBulletListPropertyMap() );

	const OUString sParaLineSpacing( CREATE_OUSTRING( "ParaLineSpacing" ) );
	//OSL_TRACE( "OOX: ParaLineSpacing unit = %d, value = %d", maLineSpacing.nUnit, maLineSpacing.nValue );
	rPropertyMap[ sParaLineSpacing ] <<= maLineSpacing.toLineSpacing();

	const OUString sParaTopMargin( CREATE_OUSTRING( "ParaTopMargin" ) );
	//OSL_TRACE( "OOX: ParaTopMargin unit = %d, value = %d", maSpaceBefore.nUnit, maSpaceBefore.nValue );
	rPropertyMap[ sParaTopMargin ] <<= maSpaceBefore.nValue;	// TODO 1/100mm

	const OUString sParaBottomMargin( CREATE_OUSTRING( "ParaBottomMargin" ) );
	//OSL_TRACE( "OOX: ParaBottomMargin unit = %d, value = %d", maSpaceAfter.nUnit, maSpaceAfter.nValue );
	rPropertyMap[ sParaBottomMargin ] <<= maSpaceAfter.nValue; // TODO 1/100mm

	::std::list< TabStop >::size_type nTabCount = maTabList.size();
	if( nTabCount != 0 )
	{
		Sequence< TabStop > aSeq( nTabCount );
		TabStop * aArray = aSeq.getArray();
		OSL_ENSURE( aArray != NULL, "sequence array is NULL" );
		::std::copy( maTabList.begin(), maTabList.end(), aArray );
		const OUString sParaTabStops( CREATE_OUSTRING( "ParaTabStops" ) );
		rPropertyMap[ sParaTabStops ] <<= aSeq;
	}

	if( maBulletListProps.is() )
	{
		const rtl::OUString sIsNumbering( CREATE_OUSTRING( "IsNumbering" ) );
		rPropertyMap[ sIsNumbering ] <<= sal_True;
		maBulletListProps.pushToProperties( rBulletListPropertyMap );
	}
	sal_Int16 nLevel = mpTextParagraphPropertiesPtr->getLevel();
	OSL_TRACE("OOX: numbering level = %d", nLevel );
	const OUString sNumberingLevel( CREATE_OUSTRING( "NumberingLevel" ) );
	rPropertyMap[ sNumberingLevel ] <<= (sal_Int16)nLevel;
	sal_Bool bTmp = sal_True;
	const OUString sNumberingIsNumber( CREATE_OUSTRING( "NumberingIsNumber" ) );
	rPropertyMap[ sNumberingIsNumber ] <<= bTmp;
}

// --------------------------------------------------------------------

void TextParagraphPropertiesContext::endFastElement( sal_Int32 ) throw (SAXException, RuntimeException)
{
}



// --------------------------------------------------------------------

Reference< XFastContextHandler > TextParagraphPropertiesContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& rXAttributes ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;
	switch( aElementToken )
	{
		case NMSP_DRAWINGML|XML_lnSpc:			// CT_TextSpacing
			xRet.set( new TextSpacingContext( getHandler(), maLineSpacing ) );
			break;
		case NMSP_DRAWINGML|XML_spcBef:			// CT_TextSpacing
			xRet.set( new TextSpacingContext( getHandler(), maSpaceBefore ) );
			break;
		case NMSP_DRAWINGML|XML_spcAft:			// CT_TextSpacing
			xRet.set( new TextSpacingContext( getHandler(), maSpaceAfter ) );
			break;

		// EG_TextBulletColor
		case NMSP_DRAWINGML|XML_buClrTx:		// CT_TextBulletColorFollowText ???
			maBulletListProps.mbBulletColorFollowText = true;
			break;
		case NMSP_DRAWINGML|XML_buClr:			// CT_Color
			xRet.set( new colorChoiceContext( getHandler(), maBulletListProps.mnBulletColor ) );
			maBulletListProps.mbHasBulletColor = true;
			break;

		// EG_TextBulletSize
		case NMSP_DRAWINGML|XML_buSzTx:			// CT_TextBulletSizeFollowText
			maBulletListProps.setBulletSize(100);
			break;
		case NMSP_DRAWINGML|XML_buSzPct:		// CT_TextBulletSizePercent
			maBulletListProps.setBulletSize( static_cast<sal_Int16>( GetPercent( rXAttributes->getOptionalValue( XML_val ) ) / 1000 ) );
			break;
		case NMSP_DRAWINGML|XML_buSzPts:		// CT_TextBulletSizePoint
			maBulletListProps.setBulletSize(0);
			maBulletListProps.setFontSize( static_cast<sal_Int16>(GetTextSize( rXAttributes->getOptionalValue( XML_val ) ) ) );
			break;

		// EG_TextBulletTypeface
		case NMSP_DRAWINGML|XML_buFontTx:		// CT_TextBulletTypefaceFollowText
			maBulletListProps.mbBulletFontFollowText = true;
			break;
		case NMSP_DRAWINGML|XML_buFont:			// CT_TextFont
			xRet.set( new TextFontContext( getHandler(), aElementToken, rXAttributes,
																		 maBulletListProps.maBulletFont ) );
			break;

		// EG_TextBullet
		case NMSP_DRAWINGML|XML_buNone:			// CT_TextNoBullet
			maBulletListProps.setNone();
			break;
		case NMSP_DRAWINGML|XML_buAutoNum:		// CT_TextAutonumberBullet
		{
			AttributeList attribs( rXAttributes );
			try {
				sal_Int32 nType = rXAttributes->getValueToken( XML_type );
				sal_Int32 nStartAt = attribs.getInteger( XML_startAt, 1 );
				if( nStartAt > 32767 )
				{
					nStartAt = 32767;
				}
				else if( nStartAt < 1 )
				{
					nStartAt = 1;
				}
				maBulletListProps.setStartAt( nStartAt );
				maBulletListProps.setType( nType );
			}
			catch(SAXException& /* e */ )
			{
				OSL_TRACE("OOX: SAXException in XML_buAutoNum");
			}
			break;
		}
		case NMSP_DRAWINGML|XML_buChar:			// CT_TextCharBullet
			try {
				maBulletListProps.setBulletChar( rXAttributes->getValue( XML_char ) );
			}
			catch(SAXException& /* e */)
			{
				OSL_TRACE("OOX: SAXException in XML_buChar");
			}
			break;
		case NMSP_DRAWINGML|XML_buBlip:			// CT_TextBlipBullet
			// TODO
			break;

		case NMSP_DRAWINGML|XML_tabLst:			// CT_TextTabStopList
			xRet.set( new TextTabStopListContext( this, maTabList ) );
			break;
		case NMSP_DRAWINGML|XML_defRPr:			// CT_TextCharacterProperties
			xRet.set( new TextCharacterPropertiesContext( this, rXAttributes, mpTextParagraphPropertiesPtr->getTextCharacterProperties() ) );
			break;
	}
	if ( !xRet.is() )
		xRet.set( this );
	return xRet;
}

// --------------------------------------------------------------------

} }

