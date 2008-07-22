/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: containerhelper.hxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: dr $ $Date: 2007/07/26 15:06:53 $
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

#ifndef OOX_CORE_CONTAINERHELPER_HXX
#define OOX_CORE_CONTAINERHELPER_HXX

#include <vector>
#include <map>
#include <boost/shared_ptr.hpp>
#include <boost/mem_fn.hpp>
#include <com/sun/star/uno/Reference.h>
#include <com/sun/star/uno/Sequence.h>

namespace rtl { class OUString; }

namespace com { namespace sun { namespace star {
    namespace container { class XIndexContainer; }
    namespace container { class XNameAccess; }
    namespace container { class XNameContainer; }
} } }

namespace oox {
namespace core {

// ============================================================================

/** Template for a vector of ref-counted objects with additional accessor functions.

    An instance of the class RefVector< Type > stores elements of the type
    ::boost::shared_ptr< Type >. The new accessor functions has() and get()
    work correctly for indexes out of the current range, there is no need to
    check the passed index before.
 */
template< typename ObjType >
class RefVector : public ::std::vector< ::boost::shared_ptr< ObjType > >
{
public:
    typedef ::std::vector< ::boost::shared_ptr< ObjType > > container_type;
    typedef typename container_type::value_type             value_type;
    typedef typename container_type::size_type              size_type;

public:
    /** Returns true, if the object with the passed index exists. Returns
        false, if the vector element exists but is an empty reference. */
    inline bool         has( sal_Int32 nIndex ) const
                        {
                            const value_type* pxRef = getRef( nIndex );
                            return pxRef && pxRef->get();
                        }

    /** Returns a reference to the object with the passed index, or 0 on error. */
    inline value_type   get( sal_Int32 nIndex ) const
                        {
                            if( const value_type* pxRef = getRef( nIndex ) ) return *pxRef;
                            return value_type();
                        }

    /** Returns the index of the last element, or -1, if the vector is empty.
        Does *not* check whether the last element is an empty reference. */
    inline sal_Int32    getLastIndex() const { return static_cast< sal_Int32 >( this->size() ) - 1; }

    /** Calls the passed member function of ObjType on every contained object. */
    template< typename FunctorType >
    inline void         forEach( FunctorType aFunctor ) const
                        {
                            ::std::for_each( this->begin(), this->end(), aFunctor );
                        }

    /** Calls the passed member function of ObjType on every contained object. */
    template< typename FuncType >
    inline void         forEachMem( FuncType pFunc ) const
                        {
                            ::std::for_each( this->begin(), this->end(), ::boost::mem_fn( pFunc ) );
                        }

private:
    inline const value_type* getRef( sal_Int32 nIndex ) const
                        {
                            return ((0 <= nIndex) && (static_cast< size_type >( nIndex ) < this->size())) ?
                                &(*this)[ static_cast< size_type >( nIndex ) ] : 0;
                        }
};

// ============================================================================

/** Template for a map of ref-counted objects with additional accessor functions.

    An instance of the class RefMap< Type > stores elements of the type
    ::boost::shared_ptr< Type >. The new accessor functions has() and get()
    work correctly for nonexisting keys, there is no need to check the passed
    key before.
 */
template< typename KeyType, typename ObjType >
class RefMap : public ::std::map< KeyType, ::boost::shared_ptr< ObjType > >
{
public:
    typedef ::std::map< KeyType, ::boost::shared_ptr< ObjType > >   container_type;
    typedef typename container_type::key_type                       key_type;
    typedef typename container_type::data_type                      data_type;
    typedef typename container_type::value_type                     value_type;

public:
    /** Returns true, if the object accossiated to the passed key exists.
        Returns false, if the key exists but points to an empty reference. */
    inline bool         has( key_type nKey ) const
                        {
                            const data_type* pxRef = getRef( nKey );
                            return pxRef && pxRef->get();
                        }

    /** Returns a reference to the object accossiated to the passed key, or 0 on error. */
    inline data_type    get( key_type nKey ) const
                        {
                            if( const data_type* pxRef = getRef( nKey ) ) return *pxRef;
                            return data_type();
                        }

    /** Calls the passed functor for every contained object. */
    template< typename FunctorType >
    inline void         forEach( FunctorType aFunctor ) const
                        {
                            ::std::for_each( this->begin(), this->end(), Functor< FunctorType >( aFunctor ) );
                        }

    /** Calls the passed member function of ObjType on every contained object. */
    template< typename FuncType >
    inline void         forEachMem( FuncType pFunc ) const
                        {
                            ::std::for_each( this->begin(), this->end(), MemFunctor< FuncType >( pFunc ) );
                        }

private:
    template< typename FunctorType >
    struct Functor
    {
        FunctorType         maFunctor;
        inline explicit     Functor( FunctorType aFunctor ) : maFunctor( aFunctor ) {}
        inline void         operator()( const value_type& rValue ) const { maFunctor( *rValue.second ); }
    };

    template< typename FuncType >
    struct MemFunctor
    {
        FuncType            mpFunc;
        inline explicit     MemFunctor( FuncType pFunc ) : mpFunc( pFunc ) {}
        inline void         operator()( const value_type& rValue ) const { ((*rValue.second).*mpFunc)(); }
    };

    inline const data_type* getRef( key_type nKey ) const
                        {
                            typename container_type::const_iterator aIt = find( nKey );
                            return (aIt == this->end()) ? 0 : &aIt->second;
                        }
};

// ============================================================================

/** Static helper functions for improved API container handling. */
class ContainerHelper
{
public:
    // com.sun.star.container.XIndexContainer ---------------------------------

    /** Creates a new index container object from scratch. */
    static ::com::sun::star::uno::Reference< ::com::sun::star::container::XIndexContainer >
                        createIndexContainer();

    /** Inserts an object into an indexed container.

        @param rxIndexContainer  com.sun.star.container.XIndexContainer
            interface of the indexed container.

        @param nIndex  Insertion index for the object.

        @param rObject  The object to be inserted.

        @return  True = object successfully inserted.
     */
    static bool         insertByIndex(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::container::XIndexContainer >& rxIndexContainer,
                            sal_Int32 nIndex,
                            const ::com::sun::star::uno::Any& rObject );

    // com.sun.star.container.XNameContainer ----------------------------------

    /** Creates a new name container object from scratch. */
    static ::com::sun::star::uno::Reference< ::com::sun::star::container::XNameContainer >
                        createNameContainer();

    /** Returns a name that is not used in the passed name container.

        @param rxNameAccess  com.sun.star.container.XNameAccess interface of
            the name container.

        @param rSuggestedName  Suggested name for the object.

        @return  An unused name. Will be equal to the suggested name, if not
            contained, otherwise a numerical index will be appended.
     */
    static ::rtl::OUString getUnusedName(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::container::XNameAccess >& rxNameAccess,
                            const ::rtl::OUString& rSuggestedName,
                            sal_Unicode cSeparator,
                            sal_Int32 nFirstIndexToAppend = 1 );

    /** Inserts an object into a name container.

        @param rxNameContainer  com.sun.star.container.XNameContainer interface
            of the name container.

        @param rName  Exact name for the object.

        @param rObject  The object to be inserted.

        @return  True = object successfully inserted.
     */
    static bool         insertByName(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::container::XNameContainer >& rxNameContainer,
                            const ::rtl::OUString& rName,
                            const ::com::sun::star::uno::Any& rObject );

    /** Inserts an object into a name container.

        The function will use an unused name to insert the object, based on the
        suggested object name. It is possible to specify whether the existing
        object or the new inserted object will be renamed, if the container
        already has an object with the name suggested for the new object.

        @param rxNameContainer  com.sun.star.container.XNameContainer interface
            of the name container.

        @param rObject  The object to be inserted.

        @param rSuggestedName  Suggested name for the object.

        @param bRenameOldExisting  Specifies behaviour if an object with the
            suggested name already exists. If false (default), the new object
            will be inserted with a name not yet extant in the container (this
            is done by appending a numerical index to the suggested name). If
            true, the existing object will be removed and inserted with an
            unused name, and the new object will be inserted with the suggested
            name.

        @return  The final name the object is inserted with. Will always be
            equal to the suggested name, if parameter bRenameOldExisting is
            true.
     */
    static ::rtl::OUString insertByUnusedName(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::container::XNameContainer >& rxNameContainer,
                            const ::com::sun::star::uno::Any& rObject,
                            const ::rtl::OUString& rSuggestedName,
                            sal_Unicode cSeparator,
                            bool bRenameOldExisting = false );

    // std::vector ------------------------------------------------------------

    /** Creates a UNO sequence from a std::vector with copies of all elements.

        @param rVector  The vector to be converted to a sequence.

        @return  A com.sun.star.uno.Sequence object with copies of all objects
            contained in the passed vector.
     */
    template< typename Type >
    static ::com::sun::star::uno::Sequence< Type >
                            vectorToSequence( const ::std::vector< Type >& rVector );
};

// ----------------------------------------------------------------------------

template< typename Type >
::com::sun::star::uno::Sequence< Type > ContainerHelper::vectorToSequence( const ::std::vector< Type >& rVector )
{
    if( rVector.empty() )
        return ::com::sun::star::uno::Sequence< Type >();
    return ::com::sun::star::uno::Sequence< Type >( &rVector.front(), static_cast< sal_Int32 >( rVector.size() ) );
}

// ============================================================================

} // namespace core
} // namespace oox

#endif
