/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: customshapegeometry.cxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: dr $ $Date: 2007/03/29 09:53:55 $
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

#include <com/sun/star/xml/sax/FastToken.hpp>

#include "oox/drawingml/customshapegeometry.hxx"
#include "oox/core/propertymap.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using ::rtl::OUString;
using ::com::sun::star::beans::NamedValue;
using namespace ::oox::core;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;

namespace oox { namespace drawingml {

// ---------------------------------------------------------------------

static OUString GetCustomShapeType( sal_Int32 nType )
{
	OUString sType;

		//todo move somewhere general
	static struct
	{
		sal_Int32 mnToken;
		const sal_Char*	mpType;
	}
	gShapeTypes[] =
	{
		{ XML_line,	"mso-spt20" },
		{ XML_triangle, "isosceles-triangle" },
		{ XML_rtTriangle, "right-triangle" },
		{ XML_rect, "rectangle" },
		{ XML_diamond, "diamond" },
		{ XML_parallelogram, "parallelogram" },
		{ XML_trapezoid, "trapezoid" },
		{ XML_pentagon, "pentagon" },
		{ XML_hexagon, "hexagon" },
//todo		{ XML_heptagon, },
		{ XML_octagon, "octagon" },
//todo		{ XML_decagon, },
//todo		{ XML_dodecagon, },
		{ XML_star4, "star4" },
		{ XML_star5, "star5" },
//todo		{ XML_star6, },
//todo		{ XML_star7, },
		{ XML_star8, "star8" },
//todo		{ XML_star10, },
//todo		{ XML_star12, },
//todo		{ XML_star16, },
		{ XML_star24, "star24" },
//todo		{ XML_star32, },
		{ XML_roundRect, "round-rectangle" },
//todo		{ XML_round1Rect, },
//todo		{ XML_round2SameRect, },
//todo		{ XML_round2DiagRect, },
//todo		{ XML_snipRoundRect, },
//todo		{ XML_snip1Rect, },
//todo		{ XML_snip2SameRect, },
//todo		{ XML_snip2DiagRect, },
		{ XML_plaque, "mso-spt21" },
		{ XML_ellipse, "ellipse" },
//todo		{ XML_teardrop, },
		{ XML_homePlate, "pentagon-right" },
		{ XML_chevron, "chevron" },
//todo		{ XML_pieWedge, },
		{ XML_blockArc, "block-arc" },
		{ XML_donut, "ring" },
		{ XML_noSmoking, "forbidden" },
		{ XML_rightArrow, "right-arrow" },
		{ XML_leftArrow, "left-arrow" },
		{ XML_upArrow, "up-arrow" },
		{ XML_downArrow, "down-arrow" },
		{ XML_stripedRightArrow, "striped-right-arrow" },
		{ XML_notchedRightArrow, "notched-right-arrow"},
		{ XML_bentUpArrow, "mso-spt90" },
		{ XML_leftRightArrow, "left-right-arrow" },
		{ XML_upDownArrow, "up-down-arrow" },
		{ XML_leftUpArrow, "mso-spt89" },
		{ XML_leftRightUpArrow, "mso-spt182" },
		{ XML_quadArrow, "quad-arrow" },
		{ XML_leftArrowCallout, "left-arrow-callout" },
		{ XML_rightArrowCallout, "right-arrow-callout" },
		{ XML_upArrowCallout, "up-arrow-callout" },
		{ XML_downArrowCallout, "down-arrow-callout" },
		{ XML_leftRightArrowCallout, "left-right-arrow-callout" },
		{ XML_upDownArrowCallout, "up-down-arrow-callout" },
		{ XML_quadArrowCallout, "quad-arrow-callout" },
		{ XML_bentArrow, "mso-spt91" },
		{ XML_uturnArrow, "mso-spt101" },
		{ XML_circularArrow, "circular-arrow" },
//todo		{ XML_leftCircularArrow, },
//todo		{ XML_leftRightCircularArrow, },
		{ XML_curvedRightArrow, "mso-spt102" },
		{ XML_curvedLeftArrow, "mso-spt103" },
		{ XML_curvedUpArrow, "mso-spt104" },
		{ XML_curvedDownArrow, "mso-spt105" },
//todo		{ XML_swooshArrow, },
		{ XML_cube, "cube" },
		{ XML_can, "can" },
		{ XML_lightningBolt, "lightning" },
		{ XML_heart, "heart" },
		{ XML_sun, "sun" },
		{ XML_moon, "moon" },
		{ XML_smileyFace, "smiley" },
		{ XML_irregularSeal1, "mso-spt71" },
		{ XML_irregularSeal2, "bang" },
		{ XML_foldedCorner, "paper" },
		{ XML_bevel, "quad-bevel" },
		{ XML_frame, "mso-spt75", },
//todo		{ XML_halfFrame, },
//todo		{ XML_corner, },
//todo		{ XML_diagStripe, },
//todo		{ XML_chord, },
		{ XML_arc, "mso-spt19" },
		{ XML_leftBracket, "left-bracket" },
		{ XML_rightBracket, "right-bracket" },
		{ XML_leftBrace, "left-brace" },
		{ XML_rightBrace, "right-brace" },
		{ XML_bracketPair, "bracket-pair" },
		{ XML_bracePair, "brace-pair" },
		{ XML_straightConnector1, "mso-spt32" },
		{ XML_bentConnector2, "mso-spt33" },
		{ XML_bentConnector3, "mso-spt34" },
		{ XML_bentConnector4, "mso-spt35" },
		{ XML_bentConnector5, "mso-spt36" },
		{ XML_curvedConnector2, "mso-spt37" },
		{ XML_curvedConnector3, "mso-spt38" },
		{ XML_curvedConnector4, "mso-spt39" },
		{ XML_curvedConnector5, "mso-spt40" },
		{ XML_callout1, "mso-spt41" },
		{ XML_callout2, "mso-spt42" },
		{ XML_callout3, "mso-spt43" },
		{ XML_accentCallout1, "mso-spt44" },
		{ XML_accentCallout2, "mso-spt45" },
		{ XML_accentCallout3, "mso-spt46" },
		{ XML_borderCallout1, "line-callout-1" },
		{ XML_borderCallout2, "line-callout-2" },
		{ XML_borderCallout3, "mso-spt49" },
		{ XML_accentBorderCallout1, "mso-spt50" },
		{ XML_accentBorderCallout2, "mso-spt51" },
		{ XML_accentBorderCallout3, "mso-spt52" },
		{ XML_wedgeRectCallout, "rectangular-callout" },
		{ XML_wedgeRoundRectCallout, "round-rectangular-callout" },
		{ XML_wedgeEllipseCallout, "round-callout" },
		{ XML_cloudCallout, "cloud-callout" },
		{ XML_ribbon, "mso-spt53" },
		{ XML_ribbon2, "mso-spt54" },
		{ XML_ellipseRibbon, "mso-spt107" },
		{ XML_ellipseRibbon2, "mso-spt108" },
		{ XML_verticalScroll, "vertical-scroll" },
		{ XML_horizontalScroll, "horizontal-scroll" },
		{ XML_wave, "mso-spt64" },
		{ XML_doubleWave, "mso-spt188" },
		{ XML_plus, "cross" },
		{ XML_flowChartProcess, "flowchart-process" },
		{ XML_flowChartDecision, "flowchart-decision" },
		{ XML_flowChartInputOutput, "flowchart-data" },
		{ XML_flowChartPredefinedProcess, "flowchart-predefined-process" },
		{ XML_flowChartInternalStorage, "flowchart-internal-storage" },
		{ XML_flowChartDocument, "flowchart-document" },
		{ XML_flowChartMultidocument, "flowchart-multidocument" },
		{ XML_flowChartTerminator, "flowchart-terminator" },
		{ XML_flowChartPreparation, "flowchart-preparation" },
		{ XML_flowChartManualInput, "flowchart-manual-input" },
		{ XML_flowChartManualOperation, "flowchart-manual-operation" },
		{ XML_flowChartConnector, "flowchart-connector" },
		{ XML_flowChartPunchedCard, "flowchart-card" },
		{ XML_flowChartPunchedTape, "flowchart-punched-tape" },
		{ XML_flowChartSummingJunction, "flowchart-summing-junction" },
		{ XML_flowChartOr, "flowchart-or" },
		{ XML_flowChartCollate, "flowchart-collate" },
		{ XML_flowChartSort, "flowchart-sort" },
		{ XML_flowChartExtract, "flowchart-extract" },
		{ XML_flowChartMerge, "flowchart-merge" },
		{ XML_flowChartOfflineStorage, "mso-spt129" },
		{ XML_flowChartOnlineStorage, "flowchart-stored-data" },
		{ XML_flowChartMagneticTape, "flowchart-sequential-access" },
		{ XML_flowChartMagneticDisk, "flowchart-magnetic-disk" },
		{ XML_flowChartMagneticDrum, "flowchart-direct-access-storage" },
		{ XML_flowChartDisplay, "flowchart-display" },
		{ XML_flowChartDelay, "flowchart-delay" },
		{ XML_flowChartAlternateProcess, "flowchart-alternate-process" },
		{ XML_flowChartOffpageConnector, "flowchart-off-page-connector" },
		{ XML_actionButtonBlank, "mso-spt189" },
		{ XML_actionButtonHome, "mso-spt190" },
		{ XML_actionButtonHelp, "mso-spt191" },
		{ XML_actionButtonInformation, "mso-spt192" },
		{ XML_actionButtonForwardNext, "mso-spt193" },
		{ XML_actionButtonBackPrevious, "mso-spt194" },
		{ XML_actionButtonEnd, "mso-spt195" },
		{ XML_actionButtonBeginning, "mso-spt196" },
		{ XML_actionButtonReturn, "mso-spt197" },
		{ XML_actionButtonDocument, "mso-spt198" },
		{ XML_actionButtonSound, "mso-spt199" },
		{ XML_actionButtonMovie, "mso-spt200" },
//todo		{ XML_gear6, },
//todo		{ XML_gear9, },
//todo		{ XML_funnel, },
//todo		{ XML_mathPlus, },
//todo		{ XML_mathMinus, },
//todo		{ XML_mathMultiply, },
//todo		{ XML_mathDivide, },
//todo		{ XML_mathEqual, },
//todo		{ XML_mathNotEqual, },
//todo		{ XML_nonIsoscelesTrapezoid, },
		{ -1, 0 }
	};

	int i = 0;
	while( gShapeTypes[i].mpType && (gShapeTypes[i].mnToken != nType) ) i++;

	if( gShapeTypes[i].mpType )
		sType = OUString::createFromAscii( gShapeTypes[i].mpType );

	return sType;
}

// ---------------------------------------------------------------------

CustomShapeGeometryContext::CustomShapeGeometryContext( const FragmentHandlerRef& xHandler, sal_Int32 nElement, const Reference< XFastAttributeList >& xAttribs, PropertyMap& rProperties )
: Context( xHandler )
, mrProperties( rProperties )
{
	OUString sShapeType;

	if( nElement == (NMSP_DRAWINGML|XML_prstGeom) )
	{
		// ST_ShapeType shape to use
		sShapeType = GetCustomShapeType( xAttribs->getOptionalValueToken( XML_prst, FastToken::DONTKNOW ) );
        OSL_ENSURE( sShapeType.getLength(), "oox::drawingml::CustomShapeCustomGeometryContext::CustomShapeCustomGeometryContext(), unknown shape type" );
	}

	if( !sShapeType.getLength() )
		sShapeType = OUString( RTL_CONSTASCII_USTRINGPARAM( "non-primitive" ) );

	static const OUString sType( RTL_CONSTASCII_USTRINGPARAM( "Type" ) );
	mrProperties[ sType ] <<= sShapeType;
}

Reference< XFastContextHandler > CustomShapeGeometryContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& ) throw (SAXException, RuntimeException)
{
	switch( aElementToken )
	{
	// todo
	case NMSP_DRAWINGML|XML_avLst:		// CT_GeomGuideList adjust value list
	case NMSP_DRAWINGML|XML_gdLst:		// CT_GeomGuideList guide list
	case NMSP_DRAWINGML|XML_ahLst:		// CT_AdjustHandleList adjust handle list
	case NMSP_DRAWINGML|XML_cxnLst:	// CT_ConnectionSiteList connection site list
	case NMSP_DRAWINGML|XML_rect:	// CT_GeomRectList geometry rect list
	case NMSP_DRAWINGML|XML_pathLst:	// CT_Path2DList 2d path list
		break;
	}

	Reference< XFastContextHandler > xEmpty;
	return xEmpty;
}

} }
