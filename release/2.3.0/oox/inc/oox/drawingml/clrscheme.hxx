/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: clrscheme.hxx,v $
 *
 *  $Revision: 1.1.2.6 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/02 11:26:00 $
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

#ifndef OOX_DRAWINGML_CLRSCHEME_HXX
#define OOX_DRAWINGML_CLRSCHEME_HXX

#include <boost/shared_ptr.hpp>
#include <map>
#include <vector>
#include "oox/core/namespaces.hxx"

namespace oox { namespace drawingml {

class ClrScheme
{
	std::map < sal_Int32, sal_Int32 > maClrScheme;

public:

	ClrScheme();
    ~ClrScheme();

    sal_Bool getColor( const sal_Int32 nSchemeClrToken, sal_Int32& rColor ) const;
	void	 setColor( const sal_Int32 nSchemeClrToken, sal_Int32 nColor );

    static bool getSystemColor( const sal_Int32 nSysClrToken, sal_Int32& rColor );
};

typedef boost::shared_ptr< ClrScheme > ClrSchemePtr;

} }

#endif  //  OOX_DRAWINGML_CLRSCHEME_HXX
