/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: xmlfilterbase.hxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: sj $ $Date: 2007/09/04 17:06:15 $
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

#ifndef OOX_CORE_XMLFILTERBASE_HXX
#define OOX_CORE_XMLFILTERBASE_HXX

#include <rtl/ref.hxx>
#include "oox/vml/drawing.hxx"
#include "oox/core/filterbase.hxx"
#include "oox/core/relations.hxx"

namespace com { namespace sun { namespace star {
    namespace xml { namespace sax { class XLocator; } }
    namespace xml { namespace sax { class XFastDocumentHandler; } }
} } }

namespace oox {
namespace core {

// ============================================================================

struct XmlFilterBaseImpl;

class XmlFilterBase : public FilterBase
{
public:
    explicit            XmlFilterBase(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::lang::XMultiServiceFactory >& rxFactory );

    virtual             ~XmlFilterBase();

    /** Has to be implemented by each filter to resolve scheme colors. */
    virtual sal_Int32   getSchemeClr( sal_Int32 nColorSchemeToken ) = 0;


	virtual const oox::vml::DrawingPtr getDrawings();

    // ------------------------------------------------------------------------

    const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XLocator >&
                        getLocator() const;

    bool                importFragment(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastDocumentHandler >& xHandler,
                            const ::rtl::OUString& rFragmentPath );

    RelationsRef        getRelations( const ::rtl::OUString& rRelPath );

    /** Copies the picture element specified with rPicturePath from the source
        document to the target models picture substorage.

        @return
            The URL of the picture in the target models storage.
	*/
    ::rtl::OUString     copyPictureStream( const ::rtl::OUString& rPicturePath );

private:
    virtual StorageRef  implCreateStorage(
                            ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream >& rxInStream,
                            ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream >& rxOutStream ) const;

private:
    ::std::auto_ptr< XmlFilterBaseImpl > mxImpl;
};

typedef ::rtl::Reference< XmlFilterBase > XmlFilterRef;

// ============================================================================

} // namespace core
} // namespace oox

#endif

