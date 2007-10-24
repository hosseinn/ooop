/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: animationspersist.hxx,v $
 *
 *  $Revision: 1.1.2.7 $
 *
 *  last change: $Author: hub $ $Date: 2007/07/12 20:00:56 $
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


#ifndef OOX_PPT_ANIMATIONPERSIST
#define OOX_PPT_ANIMATIONPERSIST

#include <list>
#include <boost/shared_ptr.hpp>

#include <rtl/ustring.hxx>

#include <com/sun/star/uno/Any.hxx>
#include <com/sun/star/drawing/XShape.hpp>

#include "oox/drawingml/drawingmltypes.hxx"
#include "oox/ppt/slidepersist.hxx"

namespace oox { namespace ppt {

	enum {
		NP_TO, NP_FROM, NP_BY, NP_USERDATA, NP_ATTRIBUTENAME,
		NP_ACCELERATION, NP_AUTOREVERSE, NP_DECELERATE, NP_DURATION, NP_FILL,
		NP_REPEATCOUNT, NP_REPEATDURATION, NP_RESTART, 
		NP_DIRECTION, NP_COLORINTERPOLATION, NP_CALCMODE, NP_TRANSFORMTYPE,
		NP_PATH,
		NP_ENDSYNC, NP_ITERATETYPE, NP_ITERATEINTERVAL,
		NP_SUBITEM, NP_TARGET, NP_COMMAND, NP_PARAMETER,
		NP_VALUES, NP_FORMULA, NP_KEYTIMES
	};

	typedef std::map< sal_Int32, ::com::sun::star::uno::Any > NodePropertyMap;


	/** data for CT_TLShapeTargetElement */
	struct ShapeTargetElement
	{
		ShapeTargetElement()
			: mnType( 0 )
			{}
		void convert( ::com::sun::star::uno::Any & aAny, sal_Int16 & rSubType ) const;

		sal_Int32               mnType;
		sal_Int32               mnRangeType;
		drawingml::IndexRange   maRange;
		::rtl::OUString msSubShapeId;
	};


	/** data for CT_TLTimeTargetElement */
	struct AnimTargetElement
	{
		AnimTargetElement()
			: mnType( 0 )
			{}
		/** convert to a set of properties */
		void convert( NodePropertyMap & aProperties, const SlidePersistPtr & pSlide) const;

		sal_Int32                  mnType;
		::rtl::OUString            msValue;

	  ShapeTargetElement         maShapeTarget;
	};

	typedef boost::shared_ptr< AnimTargetElement > AnimTargetElementPtr;

	/** data for CT_TLTimeCondition */
	struct AnimationCondition
	{
		AnimationCondition()
			: mnCondition( 0 )
			{}
		sal_Int32                  mnCondition;
		sal_Int32                  mnDelay;
		sal_Int32                  mnEvent;
		::com::sun::star::uno::Any maValue;
		AnimTargetElementPtr &     getTarget()
			{ if(!mpTarget) mpTarget.reset( new AnimTargetElement ); return mpTarget; }
	private:
		AnimTargetElementPtr       mpTarget;
	};

	typedef ::std::list< AnimationCondition > AnimationConditionList;


	struct TimeAnimationValue
	{
		::rtl::OUString            msFormula;
		::rtl::OUString            msTime;
		::com::sun::star::uno::Any maValue;
	};

	typedef ::std::list< TimeAnimationValue > TimeAnimationValueList;

} } 





#endif
