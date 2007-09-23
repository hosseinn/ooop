/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: stylesfragment.cxx,v $
 *
 *  $Revision: 1.1.2.20 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 14:57:49 $
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

#include "oox/xls/stylesfragment.hxx"
#include "oox/xls/dxfscontext.hxx"


using ::com::sun::star::uno::Reference;
using ::com::sun::star::xml::sax::XFastContextHandler;
using ::oox::core::AttributeList;
using ::rtl::OUString;

namespace oox {
namespace xls {

// ============================================================================

OoxStylesFragment::OoxStylesFragment(
        const GlobalDataHelper& rGlobalData, const OUString& rFragmentPath ) :
    GlobalFragmentBase( rGlobalData, rFragmentPath ),
    mnGradIndex( -1 )
{
}

// oox.xls.ContextHelper interface --------------------------------------------

bool OoxStylesFragment::onCanCreateContext( sal_Int32 nElement )
{
    sal_Int32 nCurrContext = getCurrentContext();
    switch( nCurrContext )
    {
        case XML_ROOT_CONTEXT:
            return  (nElement == XLS_TOKEN( styleSheet ));
        case XLS_TOKEN( styleSheet ):
            return  (nElement == XLS_TOKEN( colors )) ||
                    (nElement == XLS_TOKEN( fonts )) ||
                    (nElement == XLS_TOKEN( numFmts )) ||
                    (nElement == XLS_TOKEN( borders )) ||
                    (nElement == XLS_TOKEN( fills )) ||
                    (nElement == XLS_TOKEN( cellXfs )) ||
                    (nElement == XLS_TOKEN( dxfs )) ||
                    (nElement == XLS_TOKEN( cellStyleXfs )) ||
                    (nElement == XLS_TOKEN( dxfs )) ||
                    (nElement == XLS_TOKEN( cellStyles ));

        case XLS_TOKEN( colors ):
        case XLS_TOKEN( indexedColors ):
            return ColorPalette::isSupportedContext( nElement, nCurrContext );

        case XLS_TOKEN( fonts ):
        case XLS_TOKEN( font ):
            return Font::isSupportedContext( nElement, nCurrContext );

        case XLS_TOKEN( numFmts ):
            return  (nElement == XLS_TOKEN( numFmt ));

        case XLS_TOKEN( borders ):
        case XLS_TOKEN( border ):
        case XLS_TOKEN( left ):
        case XLS_TOKEN( right ):
        case XLS_TOKEN( top ):
        case XLS_TOKEN( bottom ):
        case XLS_TOKEN( diagonal ):
            return Border::isSupportedContext( nElement, nCurrContext );

        case XLS_TOKEN( fills ):
        case XLS_TOKEN( fill ):
        case XLS_TOKEN( patternFill ):
        case XLS_TOKEN( gradientFill ):
        case XLS_TOKEN( stop ):
            return Fill::isSupportedContext( nElement, nCurrContext );

        case XLS_TOKEN( cellStyleXfs ):
        case XLS_TOKEN( cellXfs ):
        case XLS_TOKEN( xf ):
            return Xf::isSupportedContext( nElement, nCurrContext );

        case XLS_TOKEN( cellStyles ):
            return CellStyle::isSupportedContext( nElement, nCurrContext );
    }
    return false;
}

Reference< XFastContextHandler > OoxStylesFragment::onCreateContext( sal_Int32 nElement, const AttributeList& /*rAttribs*/ )
{
    switch( nElement )
    {
        case XLS_TOKEN( dxfs ):
            return new OoxDxfsContext( *this );
    }
    return this;
}

void OoxStylesFragment::onStartElement( const AttributeList& rAttribs )
{
    sal_Int32 nCurrContext = getCurrentContext();
    sal_Int32 nPrevContext = getPreviousContext();

    switch( nCurrContext )
    {
        case XLS_TOKEN( color ):
            switch( nPrevContext )
            {
                case XLS_TOKEN( font ):
                    if( mxFont.get() ) mxFont->importAttribs( nCurrContext, rAttribs );
                break;
                case XLS_TOKEN( stop ):
                    if( mxFill.get() ) mxFill->importColor( rAttribs, mnGradIndex );
                break;
                default:
                    if( mxBorder.get() ) mxBorder->importColor( nPrevContext, rAttribs );
            }
        break;
        case XLS_TOKEN( rgbColor ):
            getStyles().importPaletteColor( rAttribs );
        break;

        case XLS_TOKEN( font ):
            mxFont = getStyles().importFont( rAttribs );
        break;
        case XLS_TOKEN( numFmt ):
            getStyles().importNumFmt( rAttribs );
        break;
        case XLS_TOKEN( border ):
            mxBorder = getStyles().importBorder( rAttribs );
        break;

        case XLS_TOKEN( fill ):
            mxFill = getStyles().importFill( rAttribs );
        break;
        case XLS_TOKEN( patternFill ):
            if( mxFill.get() ) mxFill->importPatternFill( rAttribs, true );
        break;
        case XLS_TOKEN( fgColor ):
            if( mxFill.get() ) mxFill->importFgColor( rAttribs );
        break;
        case XLS_TOKEN( bgColor ):
            if( mxFill.get() ) mxFill->importBgColor( rAttribs );
        break;
        case XLS_TOKEN( gradientFill ):
            if( mxFill.get() ) mxFill->importGradientFill( rAttribs );
        break;
        case XLS_TOKEN( stop ):
            mnGradIndex = rAttribs.getInteger( XML_position, -1 );
        break;

        case XLS_TOKEN( xf ):
            mxXf = getStyles().importXf( nPrevContext, rAttribs );
        break;
        case XLS_TOKEN( protection ):
            if( mxXf.get() ) mxXf->importProtection( rAttribs );
        break;
        case XLS_TOKEN( alignment ):
            if( mxXf.get() ) mxXf->importAlignment( rAttribs );
        break;

        case XLS_TOKEN( cellStyle ):
            getStyles().importCellStyle( rAttribs );
        break;

        default:
            switch( nPrevContext )
            {
                case XLS_TOKEN( border ):
                    if( mxBorder.get() ) mxBorder->importStyle( nCurrContext, rAttribs );
                break;
                case XLS_TOKEN( font ):
                    if( mxFont.get() ) mxFont->importAttribs( nCurrContext, rAttribs );
                break;
            }
    }
}

void OoxStylesFragment::onEndElement( const OUString& /*rChars*/ )
{
    switch( getCurrentContext() )
    {
        case XLS_TOKEN( styleSheet ):
            getStyles().finalizeImport();
        break;
    }
}

// ============================================================================

} // namespace xls
} // namespace oox

