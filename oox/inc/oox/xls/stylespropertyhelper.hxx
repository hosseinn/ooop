/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: stylespropertyhelper.hxx,v $
 *
 *  $Revision: 1.1.2.8 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/14 13:35:18 $
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

#ifndef OOX_XLS_STYLESPROPERTYHELPER_HXX
#define OOX_XLS_STYLESPROPERTYHELPER_HXX

#include "oox/core/propertysequence.hxx"
#include <com/sun/star/awt/FontDescriptor.hpp>
#include <com/sun/star/util/CellProtection.hpp>
#include <com/sun/star/table/CellHoriJustify.hpp>
#include <com/sun/star/table/CellVertJustify.hpp>
#include <com/sun/star/table/CellOrientation.hpp>
#include <com/sun/star/table/TableBorder.hpp>
#include "oox/xls/globaldatahelper.hxx"

namespace oox { namespace core {
    class PropertySet;
} }

namespace oox {
namespace xls {

// ============================================================================

const sal_Int32 API_RGB_TRANSPARENT         = -1;       /// Transparent color for API calls.

const sal_Int16 API_LINE_NONE               = 0;
const sal_Int16 API_LINE_HAIR               = 2;
const sal_Int16 API_LINE_THIN               = 35;
const sal_Int16 API_LINE_MEDIUM             = 88;
const sal_Int16 API_LINE_THICK              = 141;

const sal_Int16 API_ESCAPE_NONE             = 0;        /// No escapement.
const sal_Int16 API_ESCAPE_SUPERSCRIPT      = 101;      /// Superscript: raise characters automatically (magic value 101).
const sal_Int16 API_ESCAPE_SUBSCRIPT        = -101;     /// Subscript: lower characters automatically (magic value -101).

const sal_Int8 API_ESCAPEHEIGHT_NONE        = 100;      /// Relative character height if not escaped.
const sal_Int8 API_ESCAPEHEIGHT_DEFAULT     = 58;       /// Relative character height if escaped.

// ============================================================================

/** Contains all API font attributes. */
struct ApiFontData
{
    typedef ::com::sun::star::awt::FontDescriptor ApiFontDescriptor;

    ApiFontDescriptor   maDesc;             /// Font descriptor holding most font information (height in twips, weight in %).
    sal_Int32           mnColor;            /// Font color.
    sal_Int16           mnEscapement;       /// Escapement style.
    sal_Int8            mnEscapeHeight;     /// Escapement font height.
    bool                mbOutline;          /// True = outlined characters.
    bool                mbShadow;           /// True = shadowed chgaracters.
    bool                mbHasWstrn;         /// True = font contains Western script characters.
    bool                mbHasAsian;         /// True = font contains Asian script characters.
    bool                mbHasCmplx;         /// True = font contains Complex script characters.

    explicit            ApiFontData();
};

// ============================================================================

/** Contains all API number format attributes. */
struct ApiNumFmtData
{
    sal_Int32           mnIndex;            /// API number format index.

    explicit            ApiNumFmtData();
};

// ============================================================================

/** Contains all API cell alignment attributes. */
struct ApiAlignmentData
{
    typedef ::com::sun::star::table::CellHoriJustify ApiCellHoriJustify;
    typedef ::com::sun::star::table::CellVertJustify ApiCellVertJustify;
    typedef ::com::sun::star::table::CellOrientation ApiCellOrientation;

    ApiCellHoriJustify  meHorJustify;       /// Horizontal alignment.
    ApiCellVertJustify  meVerJustify;       /// Vertical alignment.
    ApiCellOrientation  meOrientation;      /// Normal or stacked text.
    sal_Int32           mnRotation;         /// Text rotation angle.
    sal_Int16           mnWritingMode;      /// CTL text direction.
    sal_Int16           mnIndent;           /// Indentation.
    bool                mbWrapText;         /// True = multi-line text.
    bool                mbShrink;           /// True = shrink to fit cell size.

    explicit            ApiAlignmentData();
};

bool operator==( const ApiAlignmentData& rLeft, const ApiAlignmentData& rRight );

// ============================================================================

/** Contains all API cell protection attributes. */
struct ApiProtectionData
{
    typedef ::com::sun::star::util::CellProtection ApiCellProtection;

    ApiCellProtection   maCellProt;

    explicit            ApiProtectionData();
};

bool operator==( const ApiProtectionData& rLeft, const ApiProtectionData& rRight );

// ============================================================================

/** Contains API attributes of a complete cell border. */
struct ApiBorderData
{
    typedef ::com::sun::star::table::TableBorder    ApiTableBorder;
    typedef ::com::sun::star::table::BorderLine     ApiBorderLine;

    ApiTableBorder      maBorder;           /// Left/right/top/bottom line format.
    ApiBorderLine       maTLtoBR;           /// Diagonal top-left to bottom-right line format.
    ApiBorderLine       maBLtoTR;           /// Diagonal bottom-left to top-right line format.
    bool                mbLeftUsed;         /// True = left line format used.
    bool                mbRightUsed;        /// True = right line format used.
    bool                mbTopUsed;          /// True = top line format used.
    bool                mbBottomUsed;       /// True = bottom line format used.
    bool                mbDiagUsed;         /// True = diagonal line format used.

    explicit            ApiBorderData();
};

// ============================================================================

/** Contains API fill attributes. */
struct ApiSolidFillData
{
    sal_Int32           mnColor;            /// Fill color.
    bool                mbTransparent;      /// True = transparent area.
    bool                mbUsed;             /// True = fill data is valid.

    explicit            ApiSolidFillData();
};

// ============================================================================

/** Enumerates different types of API font property sets. */
enum FontPropertyType
{
    FONT_PROPTYPE_CELL,             /// Font properties in a spreadsheet cell.
    FONT_PROPTYPE_RICHTEXT          /// Font properties in a rich text cell or header/footer.
};

// ============================================================================

/** Helper for all styles related properties.

    This helper contains helper functions to write different cell formatting
    attributes to property sets.
 */
class StylesPropertyHelper : public GlobalDataHelper
{
public:
    explicit            StylesPropertyHelper( const GlobalDataHelper& rGlobalData );

    /** Writes font properties to the passed property set. */
    void                writeFontProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiFontData& rFontData,
                            FontPropertyType ePropType );

    /** Writes number format properties to the passed property set. */
    void                writeNumFmtProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiNumFmtData& rNumFmtData );

    /** Writes cell alignment properties to the passed property set. */
    void                writeAlignmentProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiAlignmentData& rAlignData );

    /** Writes cell propection properties to the passed property set. */
    void                writeProtectionProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiProtectionData& rProtData );

    /** Writes cell border properties to the passed property set. */
    void                writeBorderProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiBorderData& rBorderData );

    /** Writes cell fill properties to the passed property set. */
    void                writeSolidFillProperties(
                            ::oox::core::PropertySet& rPropSet,
                            const ApiSolidFillData& rFillData );

private:
    ::oox::core::PropertySequence maWstrnFontNameProps;
    ::oox::core::PropertySequence maAsianFontNameProps;
    ::oox::core::PropertySequence maCmplxFontNameProps;
    ::oox::core::PropertySequence maOtherFontProps;
    ::oox::core::PropertySequence maFontEscapeProps;
    ::oox::core::PropertySequence maAlignProps;
    ::oox::core::PropertySequence maDiagBorderProps;
    ::oox::core::PropertySequence maSolidFillProps;
    ::rtl::OUString     maNumFmtProp;
    ::rtl::OUString     maCellProtProp;
    ::rtl::OUString     maBorderProp;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

