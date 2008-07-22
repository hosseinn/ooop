/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: slidetransitioncontext.cxx,v $
 *
 *  $Revision: 1.1.2.15 $
 *
 *  last change: $Author: hub $ $Date: 2007/07/05 15:18:40 $
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

#include "comphelper/anytostring.hxx"
#include "cppuhelper/exc_hlp.hxx"

#include <com/sun/star/beans/XMultiPropertySet.hpp>
#include <com/sun/star/container/XNamed.hpp>

#include <oox/ppt/backgroundproperties.hxx>
#include "oox/ppt/slidefragmenthandler.hxx"
#include "oox/ppt/slidetransitioncontext.hxx"
#include "oox/ppt/soundactioncontext.hxx"
#include "oox/drawingml/shapegroupcontext.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"
#include "tokens.hxx"

using rtl::OUString;
using namespace ::com::sun::star;
using namespace ::oox::core;
using namespace ::oox::drawingml;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace ::com::sun::star::container;

namespace oox { namespace ppt {


SlideTransitionContext::SlideTransitionContext( const FragmentHandlerRef& xHandler, const Reference< XFastAttributeList >& xAttribs, ::oox::core::PropertyMap & aProperties ) throw()
: Context( xHandler )
, maSlideProperties( aProperties )
, mbHasTransition( sal_False )
{
	AttributeList attribs(xAttribs);

	// ST_TransitionSpeed
	maTransition.setOoxTransitionSpeed( xAttribs->getOptionalValueToken( XML_spd, XML_fast ) );

	// TODO
	attribs.getBool( XML_advClick, true );
	// TODO
	// careful. if missing, no auto advance... 0 looks like a valid value
  // for auto advance
	xAttribs->getOptionalValue( XML_advTm );

}

SlideTransitionContext::~SlideTransitionContext() throw()
{

}

Reference< XFastContextHandler > SlideTransitionContext::createFastChildContext( sal_Int32 aElementToken, const Reference< XFastAttributeList >& xAttribs ) throw (SAXException, RuntimeException)
{
	Reference< XFastContextHandler > xRet;

	switch( aElementToken )
	{
	case NMSP_PPT|XML_blinds:
	case NMSP_PPT|XML_checker:
	case NMSP_PPT|XML_comb:
	case NMSP_PPT|XML_randomBar:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, xAttribs->getOptionalValueToken( XML_dir, XML_horz ), 0);
			// ST_Direction { XML_horz, XML_vert }
		}
		break;
	case NMSP_PPT|XML_cover:
	case NMSP_PPT|XML_pull:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, xAttribs->getOptionalValueToken( XML_dir, XML_l ), 0 );
			// ST_TransitionEightDirectionType { ST_TransitionSideDirectionType {
			//                                   XML_d, XML_d, XML_r, XML_u },
			//                                   ST_TransitionCornerDirectionType {
			//                                   XML_ld, XML_lu, XML_rd, XML_ru }
		}
		break;
	case NMSP_PPT|XML_cut:
	case NMSP_PPT|XML_fade:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			AttributeList attribs(xAttribs);
			// CT_OptionalBlackTransition xdb:bool
			maTransition.setOoxTransitionType( aElementToken, attribs.getBool( XML_thruBlk, false ), 0);
		}
		break;
	case NMSP_PPT|XML_push:
	case NMSP_PPT|XML_wipe:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, xAttribs->getOptionalValueToken( XML_dir, XML_l ), 0 );
			// ST_TransitionSideDirectionType { XML_d, XML_l, XML_r, XML_u }
		}
		break;
	case NMSP_PPT|XML_split:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, xAttribs->getOptionalValueToken( XML_orient, XML_horz ),	xAttribs->getOptionalValueToken( XML_dir, XML_out ) );
			// ST_Direction { XML_horz, XML_vert }
			// ST_TransitionInOutDirectionType { XML_out, XML_in }
		}
		break;
	case NMSP_PPT|XML_zoom:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, xAttribs->getOptionalValueToken( XML_dir, XML_out ), 0 );
			// ST_TransitionInOutDirectionType { XML_out, XML_in }
		}
		break;
	case NMSP_PPT|XML_wheel:
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			AttributeList attribs(xAttribs);
			maTransition.setOoxTransitionType( aElementToken, attribs.getUnsignedInteger( XML_spokes, 4 ), 0 );
			// unsignedInt
		}
		break;
	case NMSP_PPT|XML_circle:
	case NMSP_PPT|XML_diamond:
	case NMSP_PPT|XML_dissolve:
	case NMSP_PPT|XML_newsflash:
	case NMSP_PPT|XML_plus:
	case NMSP_PPT|XML_random:
	case NMSP_PPT|XML_wedge:
		// CT_Empty
		if (!mbHasTransition)
		{
			mbHasTransition = true;
			maTransition.setOoxTransitionType( aElementToken, 0, 0 );
		}
		break;


	case NMSP_PPT|XML_sndAc: // CT_TransitionSoundAction
		//"Sound"
		xRet.set( new SoundActionContext ( this->getHandler(), maSlideProperties ) );
		break;
	case NMSP_PPT|XML_extLst: // CT_OfficeArtExtensionList
		break;
	default:
		break;
	}

	if( !xRet.is() )
		xRet.set(this);

	return xRet;
}

void SlideTransitionContext::endFastElement( sal_Int32 aElement ) throw (::com::sun::star::xml::sax::SAXException, RuntimeException)
{
	if( aElement == (NMSP_PPT|XML_transition) )
	{
		if( mbHasTransition )
		{
			maTransition.setSlideProperties( maSlideProperties );
			mbHasTransition = false;
		}
	}
}


} }
