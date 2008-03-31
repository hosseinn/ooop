/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: vmldrawing.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: sj $ $Date: 2007/09/04 17:06:17 $
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

#include "oox/vml/drawing.hxx"
#include "oox/core/namespaces.hxx"
#include "tokens.hxx"

using namespace ::oox::core;

namespace oox { namespace vml {

Drawing::Drawing()
{
}
Drawing::~Drawing()
{
}

ShapePtr Drawing::createShapeById( const rtl::OUString sId ) const
{
	ShapePtr pRet, pRef;
	std::vector< ShapePtr >::const_iterator aIter( maShapes.begin() );
	while( aIter != maShapes.end() )
	{
		if ( (*aIter)->msId == sId )
		{
			pRef = (*aIter);
			break;
		}
		aIter++;
	}
	if ( pRef )
	{
		pRet = ShapePtr( new Shape() );
		if ( pRef->msType.getLength() )
		{
			std::vector< ShapePtr >::const_iterator aShapeTypeIter( maShapeTypes.begin() );
			while( aShapeTypeIter != maShapeTypes.end() )
			{
				if ( (*aShapeTypeIter)->msType == pRef->msType )
				{
					pRet->applyAttributes( *(*aShapeTypeIter).get() );
					break;
				}
				aShapeTypeIter++;
			}
		}
		pRet->applyAttributes( *pRef.get() );
	}
	return pRet;
}

} }