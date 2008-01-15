/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: binaryfilterbase.hxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/02 11:22:30 $
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

#ifndef OOX_CORE_BINARYFILTERBASE_HXX
#define OOX_CORE_BINARYFILTERBASE_HXX

#include <rtl/ref.hxx>
#include "oox/core/filterbase.hxx"

namespace oox {
namespace core {

// ============================================================================

class BinaryFilterBase : public FilterBase
{
public:
    explicit            BinaryFilterBase(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiServiceFactory >& rxFactory );

    virtual             ~BinaryFilterBase();

private:
    virtual StorageRef  implCreateStorage(
                            ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream >& rxInStream,
                            ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream >& rxOutStream ) const;
};

typedef ::rtl::Reference< BinaryFilterBase > BinaryFilterRef;

// ============================================================================

} // namespace core
} // namespace oox

#endif
