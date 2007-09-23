/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: theme.hxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: sj $ $Date: 2007/06/22 15:23:24 $
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

#ifndef OOX_DRAWINGML_THEME_HXX
#define OOX_DRAWINGML_THEME_HXX

#include <boost/shared_ptr.hpp>
#include "oox/core/propertymap.hxx"
#include <oox/drawingml/clrscheme.hxx>
#include <map>

namespace oox { namespace drawingml {

class Shape;

typedef boost::shared_ptr< Shape > ShapePtr;

class Theme
{

public:

	Theme();
    ~Theme();

	oox::drawingml::ClrSchemePtr			getClrScheme() const { return mpClrSchemePtr; };

	rtl::OUString&							getStyleName(){ return maStyleName; };
	std::vector< oox::core::PropertyMap	>&	getFillStyleList(){ return maFillStyleList; };
	std::vector< oox::core::PropertyMap	>&	getLineStyleList(){ return maLineStyleList; };
	std::vector< oox::core::PropertyMap	>&	getEffectStyleList(){ return maEffectStyleList; };
	std::vector< oox::core::PropertyMap	>&	getBgFillStyleList(){ return maBgFillStyleList; };
	
	oox::drawingml::ShapePtr				getspDef() const { return mpspDefPtr; };
	oox::drawingml::ShapePtr				getlnDef() const { return mplnDefPtr; };
	oox::drawingml::ShapePtr				gettxDef() const { return mptxDefPtr; };

private:

	rtl::OUString							maStyleName;
	std::vector< oox::core::PropertyMap	>	maFillStyleList;
	std::vector< oox::core::PropertyMap	>	maLineStyleList;
	std::vector< oox::core::PropertyMap	>	maEffectStyleList;
	std::vector< oox::core::PropertyMap >	maBgFillStyleList;

	oox::drawingml::ClrSchemePtr			mpClrSchemePtr;
	oox::drawingml::ShapePtr				mpspDefPtr;
	oox::drawingml::ShapePtr				mplnDefPtr;
	oox::drawingml::ShapePtr				mptxDefPtr;
};

typedef boost::shared_ptr< Theme > ThemePtr;

} }

#endif  //  OOX_DRAWINGML_THEME_HXX
