/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: propertyset.hxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 14:18:20 $
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

#ifndef OOX_CORE_PROPERTYSET_HXX
#define OOX_CORE_PROPERTYSET_HXX

#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/beans/XMultiPropertySet.hpp>

namespace oox {
namespace core {

// ============================================================================

/** A wrapper for a UNO property set.

    This class provides functions to silently get and set properties (without
    exceptions, without the need to check validity of the UNO property set).

    An instance is constructed with the reference to a UNO property set or any
    other interface (the constructor will query for the
    com.sun.star.beans.XPropertySet interface then). The reference to the
    property set will be kept as long as the instance of this class is alive.

    The functions getProperties() and setProperties() try to handle all passed
    values at once, using the com.sun.star.beans.XMultiPropertySet interface.
    If the implementation does not support the XMultiPropertySet interface, all
    properties are handled separately in a loop.
 */
class PropertySet
{
public:
    inline explicit     PropertySet() {}

    /** Constructs a property set wrapper with the passed UNO property set. */
    inline explicit     PropertySet(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet >& rxPropSet,
                            bool bUseMultiPropSet = true )
                                { set( rxPropSet, bUseMultiPropSet ); }

    /** Constructs a property set wrapper after querying the XPropertySet interface. */
    template< typename Type >
    inline explicit     PropertySet( const Type& rObject, bool bUseMultiPropSet = true )
                            { set( rObject, bUseMultiPropSet ); }

    /** Sets the passed UNO property set and releases the old UNO property set. */
    void                set(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet >& rxPropSet,
                            bool bUseMultiPropSet = true );

    /** Queries the passed object (interface or any) for an XPropertySet and releases the old UNO property set. */
    template< typename Type >
    inline void         set( const Type& rObject, bool bUseMultiPropSet = true )
                            { set( ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet >( rObject, ::com::sun::star::uno::UNO_QUERY ), bUseMultiPropSet ); }

    /** Returns true, if the contained XPropertySet interface is valid. */
    inline bool         is() const { return mxPropSet.is(); }

    /** Returns the contained XPropertySet interface. */
    inline ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet >
                        getXPropertySet() const { return mxPropSet; }

    // Get properties ---------------------------------------------------------

    /** Gets the specified property from the property set.
        @return  true, if the any could be filled with the property value. */
    bool                getAnyProperty(
                            ::com::sun::star::uno::Any& orValue,
                            const ::rtl::OUString& rPropName ) const;

    /** Gets the specified property from the property set.
        @return  true, if the passed variable could be filled with the property value. */
    template< typename Type >
    inline bool         getProperty(
                            Type& orValue,
                            const ::rtl::OUString& rPropName ) const;

    /** Gets the specified boolean property from the property set.
        @return  true = property contains true; false = property contains false or error occured. */
    bool                getBoolProperty( const ::rtl::OUString& rPropName ) const;

    /** Gets the specified properties from the property set. Tries to use the XMultiPropertySet interface.
        @param orValues  (out-parameter) The related property values.
        @param rPropNames  The property names. MUST be ordered alphabetically. */
    void                getProperties(
                            ::com::sun::star::uno::Sequence< ::com::sun::star::uno::Any >& orValues,
                            const ::com::sun::star::uno::Sequence< ::rtl::OUString >& rPropNames ) const;

    // Set properties ---------------------------------------------------------

    /** Puts the passed any into the property set. */
    void                setAnyProperty(
                            const ::rtl::OUString& rPropName,
                            const ::com::sun::star::uno::Any& rValue );

    /** Puts the passed value into the property set. */
    template< typename Type >
    inline void         setProperty(
                            const ::rtl::OUString& rPropName,
                            const Type& rValue );

    /** Puts the passed properties into the property set. Tries to use the XMultiPropertySet interface.
        @param rPropNames  The property names. MUST be ordered alphabetically.
        @param rValues  The related property values. */
    void                setProperties(
                            const ::com::sun::star::uno::Sequence< ::rtl::OUString >& rPropNames,
                            const ::com::sun::star::uno::Sequence< ::com::sun::star::uno::Any >& rValues );

    // ------------------------------------------------------------------------
private:
    ::com::sun::star::uno::Reference< ::com::sun::star::beans::XPropertySet >
                        mxPropSet;          /// The mandatory property set interface.
    ::com::sun::star::uno::Reference< ::com::sun::star::beans::XMultiPropertySet >
                        mxMultiPropSet;     /// The optional multi property set interface.
};

// ----------------------------------------------------------------------------

template< typename Type >
inline bool PropertySet::getProperty( Type& orValue, const ::rtl::OUString& rPropName ) const
{
    ::com::sun::star::uno::Any aAny;
    return getAnyProperty( aAny, rPropName ) && (aAny >>= orValue);
}

template< typename Type >
inline void PropertySet::setProperty( const ::rtl::OUString& rPropName, const Type& rValue )
{
    setAnyProperty( rPropName, ::com::sun::star::uno::Any( rValue ) );
}

// ============================================================================

} // namespace core
} // namespace oox

#endif
