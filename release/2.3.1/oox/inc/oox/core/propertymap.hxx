/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: propertymap.hxx,v $
 *
 *  $Revision: 1.1.2.4 $
 *
 *  last change: $Author: sj $ $Date: 2007/08/30 12:06:42 $
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

#ifndef OOX_CORE_PROPERTYMAP_HXX
#define OOX_CORE_PROPERTYMAP_HXX

#include <map>

#include <com/sun/star/beans/NamedValue.hpp>
#include <com/sun/star/beans/PropertyValue.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/uno/Sequence.hxx>

namespace oox { namespace core {

typedef std::map< ::rtl::OUString, com::sun::star::uno::Any > PropertyMapBase;

class PropertyMap : public PropertyMapBase
{
public:
	bool hasProperty( const ::rtl::OUString& rName ) const;
	const com::sun::star::uno::Any* getPropertyValue( const ::rtl::OUString& rName ) const;

	void makeSequence( ::com::sun::star::uno::Sequence< ::com::sun::star::beans::PropertyValue >& rSequence ) const;
	void makeSequence( ::com::sun::star::uno::Sequence< ::com::sun::star::beans::NamedValue >& rSequence ) const;
	void makeSequence( ::com::sun::star::uno::Sequence< ::rtl::OUString >& rNames,
					   ::com::sun::star::uno::Sequence< ::com::sun::star::uno::Any >& rValues ) const;

	::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet > makePropertySet() const;
	void dump_debug(const char *pMessage = NULL);
};

} }


#endif // OOX_CORE_PROPERTYMAP_HXX
