/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: relations.cxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: kohei $ $Date: 2007/06/01 00:38:43 $
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

#include "oox/core/relations.hxx"

namespace oox { namespace core {

// --------------------------------------------------------------------

Relation::Relation( const rtl::OUString& rsId, const rtl::OUString& rsType, const rtl::OUString& rsTarget )
: msId( rsId ), msType( rsType ), msTarget( rsTarget )
{
}

// --------------------------------------------------------------------

RelationPtr Relations::getRelationById( const rtl::OUString& rsId ) const
{
	std::vector< RelationPtr >::const_iterator aIter( begin() );
	for( ; aIter != end(); aIter++ )
	{
		if( (*aIter)->msId == rsId )
			return (*aIter);
	}

	RelationPtr aEmpty;
	return aEmpty;
}

// --------------------------------------------------------------------

RelationPtr Relations::getRelationByType( const rtl::OUString& rsType ) const
{
	std::vector< RelationPtr >::const_iterator aIter( begin() );
	for( ; aIter != end(); aIter++ )
	{
		if( (*aIter)->msType == rsType )
			return (*aIter);
	}

	RelationPtr aEmpty;
	return aEmpty;
}

void Relations::getRelationsByType( const rtl::OUString& rsType, std::vector<RelationPtr>& raRels ) const
{
    using std::vector;
    vector< RelationPtr > aRels;
    aRels.reserve(5);
    vector< RelationPtr >::const_iterator aIter( begin() ), aIterEnd( end() );
    for( ; aIter != aIterEnd; ++aIter )
    {
        if( (*aIter)->msType == rsType )
            aRels.push_back(*aIter);
    }
    std::swap( raRels, aRels );
}

// --------------------------------------------------------------------

} }
