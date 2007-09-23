/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: shape.hxx,v $
 *
 *  $Revision: 1.1.2.17 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/23 15:56:31 $
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

#ifndef OOX_DRAWINGML_SHAPE_HXX
#define OOX_DRAWINGML_SHAPE_HXX

#include "oox/drawingml/theme.hxx"
#include "oox/core/propertymap.hxx"
#include "oox/drawingml/textbody.hxx"
#include "oox/drawingml/textliststyle.hxx"

#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/drawing/XDrawPage.hpp>
#include <boost/shared_ptr.hpp>
#include <boost/enable_shared_from_this.hpp>
#include <vector>
#include <map>

namespace oox { namespace drawingml {

class Shape;

typedef boost::shared_ptr< Shape > ShapePtr;

enum ShapeStyle
{
    SHAPESTYLE_ln,
    SHAPESTYLE_fill,
    SHAPESTYLE_effect,
	SHAPESTYLE_font
};
typedef std::map< ShapeStyle, sal_Int32 > ShapeStylesColorMap;
typedef std::map< ShapeStyle, rtl::OUString > ShapeStylesIndexMap;

class Shape
	: public boost::enable_shared_from_this< Shape >
{
public:

	Shape( const sal_Char* pServiceType = NULL );
    virtual ~Shape();

	::oox::core::PropertyMap&		getShapeProperties(){ return maShapeProperties; };
	::oox::core::PropertyMap&		getCustomShapeGeometry(){ return maCustomShapeGeometry; };

	void							setPosition( com::sun::star::awt::Point nPosition ){ maPosition = nPosition; };
	void							setSize( com::sun::star::awt::Size aSize ){ maSize = aSize; };
	void							setRotation( sal_Int32 nRotation ) { mnRotation = nRotation; };
	void							setFlip( sal_Bool bFlipH, sal_Bool bFlipV ) { mbFlipH = bFlipH; mbFlipV = bFlipV; };
	void							addChild( const ShapePtr pChildPtr ) { maChilds.push_back( pChildPtr ); };
	std::vector< ShapePtr >&		getChilds() { return maChilds; };

	void							setName( const rtl::OUString& rName ) { msName = rName; };
	void							setId( const rtl::OUString& rId ) { msId = rId; };
	void							setSubType( sal_uInt32 nSubType ) { mnSubType = nSubType; };
	sal_Int32						getSubType() const { return mnSubType; };
	void							setIndex( sal_uInt32 nIndex ) { mnIndex = nIndex; };

	// setDefaults has to be called if styles are imported (OfficeXML is not storing properties having the default value)
	void							setDefaults();

	void				setTextBody(const TextBodyPtr & pTextBody);
	TextBodyPtr			getTextBody();
	TextListStylePtr	getListStyles() const { return mpTextListStyle; }


	ShapeStylesColorMap&	getShapeStylesColor(){ return maShapeStylesColorMap; };
	ShapeStylesIndexMap&	getShapeStylesIndex(){ return maShapeStylesIndexMap; };

	// addShape is creating and inserting the corresponding XShape.
	void		addShape( const ::com::sun::star::uno::Reference< ::com::sun::star::frame::XModel > &rxModel,
														const oox::drawingml::ThemePtr, std::map< ::rtl::OUString, ShapePtr > &,
														const ::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShapes >& rxShapes,
														const ::com::sun::star::awt::Rectangle* pShapeRect );
	const ::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShape > & getXShape() const { return mxShape; }

	virtual void applyShapeReference( const oox::drawingml::Shape& rReferencedShape );

protected:

	::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShape >
						createAndInsert( const rtl::OUString& rServiceName, const ::com::sun::star::uno::Reference< ::com::sun::star::frame::XModel > &rxModel,
											const oox::drawingml::ThemePtr pThemePtr, const ::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShapes >& rxShapes,
											const ::com::sun::star::awt::Rectangle* pShapeRect );
	void addChilds( Shape& rMaster, const ::com::sun::star::uno::Reference< ::com::sun::star::frame::XModel > &rxModel,
														const oox::drawingml::ThemePtr, std::map< ::rtl::OUString, ShapePtr > &,
														const ::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShapes >& rxShapes,
														const ::com::sun::star::awt::Rectangle& rClientRect );

	std::vector< ShapePtr >		maChilds;				// only used for group shapes
	TextBodyPtr					mpTextBody;
	::oox::core::PropertyMap	maShapeProperties;
	::oox::core::PropertyMap	maCustomShapeGeometry;
	TextListStylePtr			mpTextListStyle;
	::com::sun::star::uno::Reference< ::com::sun::star::drawing::XShape > mxShape;

	rtl::OUString	msServiceName;
    rtl::OUString	msName;
    rtl::OUString	msId;
    sal_uInt32		mnSubType;		// if this type is not zero, then the shape is a placeholder
    sal_uInt32		mnIndex;

	ShapeStylesColorMap	maShapeStylesColorMap;
	ShapeStylesIndexMap maShapeStylesIndexMap;

	com::sun::star::awt::Size		maSize;
	com::sun::star::awt::Point		maPosition;

private:

	void setShapeStyles( const oox::drawingml::ThemePtr pThemePtr );

	sal_Int32						mnRotation;
	sal_Bool						mbFlipH;
	sal_Bool						mbFlipV;
};

} }

#endif  //  OOX_DRAWINGML_SHAPE_HXX
