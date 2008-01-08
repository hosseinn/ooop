/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: contexthelper.hxx,v $
 *
 *  $Revision: 1.1.2.12 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:58:00 $
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

#ifndef OOX_XLS_CONTEXTHELPER_HXX
#define OOX_XLS_CONTEXTHELPER_HXX

#include <vector>
#include <rtl/ustrbuf.hxx>
#include "tokens.hxx"
#include "oox/core/namespaces.hxx"
#include "oox/core/attributelist.hxx"

#define R_TOKEN( token )            (::oox::core::NMSP_RELATIONSHIPS | XML_##token)
#define XML_TOKEN( token )          (::oox::core::NMSP_XML | XML_##token)
#define XLS_TOKEN( token )          (::oox::core::NMSP_EXCEL | XML_##token)
#define DML_TOKEN( token )          (::oox::core::NMSP_DRAWINGML | XML_##token)

const sal_Int32 XML_ROOT_CONTEXT    = XML_TOKEN_COUNT + 1;

namespace com { namespace sun { namespace star {
    namespace xml { namespace sax { class XFastContextHandler; } }
} } }

namespace oox {
namespace xls {

// ============================================================================

/** Helper class that provides a context identifier stack.

    Fragment handlers and context handlers derived from this helper class will
    track the identifiers of the current context in a stack. The idea is to use
    the same instance of a fragment handler or context handler to process
    several nested elements in a stream. For that, the abstract function
    onCreateContext() has to return 'this' for the passed element.

    Derived classes have to implement the createFastChildContext(),
    startFastElement(), characters(), and endFastElement() functions from the
    com.sun.star.xml.sax.XFastContextHandler interface by simply calling the
    implCreateChildContext(), implStartCurrentContext(), implCharacters(), and
    implEndCurrentContext() of this helper. The new abstract functions have to
    be implemented according to the elements to be processed.
 */
class ContextHelper
{
public:
    /** Will be called to decide if the passed element can be processed in the
        current context.
     */
    virtual bool        onCanCreateContext( sal_Int32 nElement ) = 0;

    /** Will be called if a new context has to be created for the passed element.

        Usually 'this' should be returned to improve performance by reusing the
        same instance to process several elements.
     */
    virtual ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        onCreateContext( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs ) = 0;

    /** Will be called if a new context element has been started.

        The current element identifier can be accessed by using
        getCurrentContext() or isCurrentContext().
     */
    virtual void        onStartElement( const ::oox::core::AttributeList& rAttribs ) = 0;

    /** Will be called if the current context element is about to be left.

        The current element identifier can be accessed by using
        getCurrentContext() or isCurrentContext().

        @param rChars  The characters collected in this element.
     */
    virtual void        onEndElement( const ::rtl::OUString& rChars ) = 0;

protected:
    explicit            ContextHelper();
    virtual             ~ContextHelper();

    /** Returns the element identifier of the current topmost context. */
    inline sal_Int32    getCurrentContext() const
                            { return maContextStack.empty() ? XML_ROOT_CONTEXT : maContextStack.back().mnElement; }

    /** Returns true, if nElement contains the current topmost context. */
    inline bool         isCurrentContext( sal_Int32 nElement ) const
                            { return getCurrentContext() == nElement; }

    /** Returns true, if either nElement1 or nElement2 contain the current topmost context. */
    inline bool         isCurrentContext( sal_Int32 nElement1, sal_Int32 nElement2 ) const
                            { return isCurrentContext( nElement1 ) || isCurrentContext( nElement2 ); }

    /** Returns the element identifier of the specified parent context. */
    sal_Int32           getPreviousContext( sal_Int32 nCountBack = 1 ) const;

    /** Returns the element identifier of the specified parent context. */
    inline sal_Int32    isPreviousContext( sal_Int32 nElement, sal_Int32 nCountBack = 1 ) const
                            { return getPreviousContext( nCountBack ) == nElement; }

    /** Must be called from createFastChildContext() in derived classes. */
    ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastContextHandler >
                        implCreateChildContext(
                            sal_Int32 nElement,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& rxAttribs );

    /** Must be called from startFastElement() in derived classes. */
    void                implStartCurrentContext(
                            sal_Int32 nElement,
                            const ::com::sun::star::uno::Reference< ::com::sun::star::xml::sax::XFastAttributeList >& rxAttribs );

    /** Must be called from characters() in derived classes. */
    void                implCharacters( const ::rtl::OUString& rChars );

    /** Must be called from endFastElement() in derived classes. */
    void                implEndCurrentContext( sal_Int32 nElement );

private:
                        ContextHelper( const ContextHelper& );
    ContextHelper&      operator=( const ContextHelper& );

    void                appendCollectedChars();

private:
    /** Information about processed context elements. */
    struct ContextInfo
    {
        sal_Int32           mnElement;      /// The element identifier.
        ::rtl::OUStringBuffer maCurrChars;  /// Collected characters from context.
        ::rtl::OUStringBuffer maFinalChars; /// Finalized (stipped) characters.
        bool                mbTrimSpaces;   /// True = trims leading/trailing spaces from text data.
    };

    ::std::vector< ContextInfo > maContextStack;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

