/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: relations.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: kohei $ $Date: 2007/06/01 00:38:42 $
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

#ifndef OOX_CORE_RELATIONS
#define OOX_CORE_RELATIONS

#include <vector>
#include <boost/shared_ptr.hpp>
#include <cppuhelper/weak.hxx>
#include <rtl/ustring.hxx>
#include <rtl/ref.hxx>

namespace oox { namespace core {

class Relation
{
public:
	rtl::OUString	msId;
	rtl::OUString	msType;
	rtl::OUString	msTarget;

	Relation( const rtl::OUString& rsId, const rtl::OUString& rsType, const rtl::OUString& rsTarget );
};

typedef boost::shared_ptr< Relation > RelationPtr;

class Relations : public std::vector< RelationPtr >, public cppu::OWeakObject
{
public:
	RelationPtr getRelationById( const rtl::OUString& rsId ) const;
	RelationPtr getRelationByType( const rtl::OUString& rsType ) const;

    /** Finds all relation entries associated with the specified type. */
    void getRelationsByType( const rtl::OUString& rsType, std::vector< RelationPtr >& raRels ) const;
};

typedef rtl::Reference< Relations > RelationsRef;

} }

#endif // OOX_CORE_RELATIONS
