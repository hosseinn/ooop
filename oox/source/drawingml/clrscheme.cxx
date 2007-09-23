/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: clrscheme.cxx,v $
 *
 *  $Revision: 1.1.2.5 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/02 11:27:14 $
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

#include "oox/drawingml/clrscheme.hxx"
#include "tokens.hxx"

namespace oox { namespace drawingml {

ClrScheme::ClrScheme()
{
}
ClrScheme::~ClrScheme()
{
}

sal_Bool ClrScheme::getColor( const sal_Int32 nSchemeClrToken, sal_Int32& rColor ) const
{
    std::map < sal_Int32, sal_Int32 >::const_iterator aIter( maClrScheme.find( nSchemeClrToken ) );
	if ( aIter != maClrScheme.end() )
		rColor = (*aIter).second;
	return aIter != maClrScheme.end();
}

void ClrScheme::setColor( const sal_Int32 nSchemeClrToken, sal_Int32 nColor )
{
	maClrScheme[ nSchemeClrToken ] = nColor;
}

// static
bool ClrScheme::getSystemColor( const sal_Int32 nSysClrToken, sal_Int32& rColor )
{
    //! TODO: get colors from system
    switch( nSysClrToken )
    {
        case XML_window:        rColor = 0xFFFFFF;  break;
        case XML_windowText:    rColor = 0x000000;  break;
        //! TODO: more colors to follow... (chapter 5.1.12.58)
        default:    return false;
    }
    return true;
}

} }
