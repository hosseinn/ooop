/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: stylesbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.35 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/30 14:11:01 $
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

#ifndef OOX_XLS_STYLESBUFFER_HXX
#define OOX_XLS_STYLESBUFFER_HXX

#include "oox/core/containerhelper.hxx"
#include "oox/xls/globaldatahelper.hxx"
#include "oox/xls/stylespropertyhelper.hxx"
#include "oox/xls/numberformatsbuffer.hxx"

namespace com { namespace sun { namespace star {
    namespace awt { struct FontDescrtiptor; }
} } }

namespace oox { namespace core {
    class AttributeList;
    class PropertySet;
} }

namespace oox {
namespace xls {

// ============================================================================

// OOX predefined color indexes (also used in BIFF3-BIFF8)
const sal_Int32 OOX_COLOR_USEROFFSET        = 0;        /// First user defined color in palette (OOX).
const sal_Int32 BIFF_COLOR_USEROFFSET       = 8;        /// First user defined color in palette (BIFF).
const sal_Int32 OOX_COLOR_WINDOWTEXT3       = 24;       /// System window text color (BIFF3-BIFF4).
const sal_Int32 OOX_COLOR_WINDOWBACK3       = 25;       /// System window background color (BIFF3-BIFF4).
const sal_Int32 OOX_COLOR_WINDOWTEXT        = 64;       /// System window text color (BIFF5+).
const sal_Int32 OOX_COLOR_WINDOWBACK        = 65;       /// System window background color (BIFF5+).
const sal_Int32 OOX_COLOR_BUTTONBACK        = 67;       /// System button background color (face color).
const sal_Int32 OOX_COLOR_CHWINDOWTEXT      = 77;       /// System window text color (BIFF8 charts).
const sal_Int32 OOX_COLOR_CHWINDOWBACK      = 78;       /// System window background color (BIFF8 charts).
const sal_Int32 OOX_COLOR_CHBORDERAUTO      = 79;       /// Automatic frame border (BIFF8 charts).
const sal_Int32 OOX_COLOR_NOTEBACK          = 80;       /// Note background color.
const sal_Int32 OOX_COLOR_NOTETEXT          = 81;       /// Note text color.
const sal_Int32 OOX_COLOR_FONTAUTO          = 0x7FFF;   /// Font auto color (system window text color).

// OOX font family (also used in BIFF)
const sal_Int32 OOX_FONTFAMILY_NONE         = 0;
const sal_Int32 OOX_FONTFAMILY_ROMAN        = 1;
const sal_Int32 OOX_FONTFAMILY_SWISS        = 2;
const sal_Int32 OOX_FONTFAMILY_MODERN       = 3;
const sal_Int32 OOX_FONTFAMILY_SCRIPT       = 4;
const sal_Int32 OOX_FONTFAMILY_DECORATIVE   = 5;

// OOX font charset (also used in BIFF)
const sal_Int32 OOX_FONTCHARSET_UNUSED      = -1;
const sal_Int32 OOX_FONTCHARSET_ANSI        = 0;

// OOX cell text direction (also used in BIFF)
const sal_Int32 OOX_XF_TEXTDIR_CONTEXT      = 0;
const sal_Int32 OOX_XF_TEXTDIR_LTR          = 1;
const sal_Int32 OOX_XF_TEXTDIR_RTL          = 2;

// OOX cell rotation (also used in BIFF)
const sal_Int32 OOX_XF_ROTATION_NONE        = 0;
const sal_Int32 OOX_XF_ROTATION_90CCW       = 90;
const sal_Int32 OOX_XF_ROTATION_90CW        = 180;
const sal_Int32 OOX_XF_ROTATION_STACKED     = 255;

// OOX cell indentation
const sal_Int32 OOX_XF_INDENT_NONE          = 0;

// OOX built-in cell styles (also used in BIFF)
const sal_Int32 OOX_STYLE_USERDEF           = -1;
const sal_Int32 OOX_STYLE_NORMAL            = 0;        /// Default cell style.
const sal_Int32 OOX_STYLE_ROWLEVEL          = 1;        /// RowLevel_x cell style.
const sal_Int32 OOX_STYLE_COLLEVEL          = 2;        /// ColLevel_x cell style.

const sal_Int32 OOX_STYLE_NOLEVEL           = -1;
const sal_Int32 OOX_STYLE_LEVELCOUNT        = 7;        /// Number of outline level styles.

// BIFF constants -------------------------------------------------------------

// BIFF predefined color indexes
const sal_uInt16 BIFF2_COLOR_BLACK          = 0;        /// Black (text) in BIFF2.
const sal_uInt16 BIFF2_COLOR_WHITE          = 1;        /// White (background) in BIFF2.

// BIFF font flags
const sal_uInt16 BIFF_FONTFLAG_NONE         = 0x0000;
const sal_uInt16 BIFF_FONTFLAG_BOLD         = 0x0001;
const sal_uInt16 BIFF_FONTFLAG_ITALIC       = 0x0002;
const sal_uInt16 BIFF_FONTFLAG_UNDERLINE    = 0x0004;
const sal_uInt16 BIFF_FONTFLAG_STRIKEOUT    = 0x0008;
const sal_uInt16 BIFF_FONTFLAG_OUTLINE      = 0x0010;
const sal_uInt16 BIFF_FONTFLAG_SHADOW       = 0x0020;

// BIFF font underline
const sal_uInt8 BIFF_FONTUNDERL_NONE        = 0;
const sal_uInt8 BIFF_FONTUNDERL_SINGLE      = 1;
const sal_uInt8 BIFF_FONTUNDERL_DOUBLE      = 2;
const sal_uInt8 BIFF_FONTUNDERL_SINGLE_ACC  = 33;
const sal_uInt8 BIFF_FONTUNDERL_DOUBLE_ACC  = 34;

// BIFF font escapement
const sal_uInt16 BIFF_FONTESC_NONE          = 0;
const sal_uInt16 BIFF_FONTESC_SUPER         = 1;
const sal_uInt16 BIFF_FONTESC_SUB           = 2;

// BIFF XF flags
const sal_uInt16 BIFF_XF_LOCKED             = 0x0001;
const sal_uInt16 BIFF_XF_HIDDEN             = 0x0002;
const sal_uInt16 BIFF_XF_STYLE              = 0x0004;
const sal_uInt16 BIFF_XF_STYLEPARENT        = 0x0FFF;   /// Syles don't have a parent.
const sal_uInt16 BIFF_XF_LINEBREAK          = 0x0008;   /// Automatic line break.
const sal_uInt16 BIFF_XF_SHRINK             = 0x0010;   /// Shrink to fit into cell.
const sal_uInt16 BIFF_XF_MERGE              = 0x0020;

// BIFF XF attribute used flags
const sal_uInt8 BIFF_XF_DIFF_VALFMT         = 0x01;
const sal_uInt8 BIFF_XF_DIFF_FONT           = 0x02;
const sal_uInt8 BIFF_XF_DIFF_ALIGN          = 0x04;
const sal_uInt8 BIFF_XF_DIFF_BORDER         = 0x08;
const sal_uInt8 BIFF_XF_DIFF_AREA           = 0x10;
const sal_uInt8 BIFF_XF_DIFF_PROT           = 0x20;

// BIFF XF text orientation
const sal_uInt8 BIFF_XF_ORIENT_NONE         = 0;
const sal_uInt8 BIFF_XF_ORIENT_STACKED      = 1;        /// Stacked top to bottom.
const sal_uInt8 BIFF_XF_ORIENT_90CCW        = 2;        /// 90 degr. counterclockwise.
const sal_uInt8 BIFF_XF_ORIENT_90CW         = 3;        /// 90 degr. clockwise.

// BIFF XF line styles
const sal_uInt8 BIFF_LINE_NONE              = 0;
const sal_uInt8 BIFF_LINE_THIN              = 1;

// BIFF XF patterns
const sal_uInt8 BIFF_PATT_NONE              = 0;
const sal_uInt8 BIFF_PATT_125               = 17;

// BIFF2 XF flags
const sal_uInt8 BIFF2_XF_VALFMT_MASK        = 0x3F;
const sal_uInt8 BIFF2_XF_LOCKED             = 0x40;
const sal_uInt8 BIFF2_XF_HIDDEN             = 0x80;
const sal_uInt8 BIFF2_XF_LEFTLINE           = 0x08;
const sal_uInt8 BIFF2_XF_RIGHTLINE          = 0x10;
const sal_uInt8 BIFF2_XF_TOPLINE            = 0x20;
const sal_uInt8 BIFF2_XF_BOTTOMLINE         = 0x40;
const sal_uInt8 BIFF2_XF_BACKGROUND         = 0x80;

// BIFF8 diagonal borders
const sal_uInt32 BIFF_XF_DIAG_TLBR          = 0x40000000;   /// Top-left to bottom-right.
const sal_uInt32 BIFF_XF_DIAG_BLTR          = 0x80000000;   /// Bottom-left to top-right.

// BIFF STYLE flags
const sal_uInt16 BIFF_STYLE_BUILTIN         = 0x8000;
const sal_uInt16 BIFF_STYLE_XFMASK          = 0x0FFF;

// ============================================================================

/** Generic helper struct for a style color. */
struct OoxColor
{
    enum Type
    {
        TYPE_RGB,       /// RGB value.
        TYPE_THEME,     /// Index to theme color.
        TYPE_PALETTE    /// Index to palette.
    };

    double              mfTint;             /// Color tint (darken/lighten).
    Type                meType;             /// Type of mnValue (see enumeration above).
    sal_Int32           mnValue;            /// RGB color value or palette index.

    explicit            OoxColor();
    explicit            OoxColor( Type eType, sal_Int32 nValue, double fTint = 0.0 );

    /** Sets the color to the passed values. */
    void                set( Type eType, sal_Int32 nValue, double fTint = 0.0 );

    /** Imports the color from the passed attribute list. */
    void                importColor( const ::oox::core::AttributeList& rAttribs );
    /** Imports a palette color identifier from the passed BIFF stream. */
    void                importColorId( BiffInputStream& rStrm, bool b16Bit = true );
    /** Imports an RGB color value from the passed BIFF stream. */
    void                importColorRgb( BiffInputStream& rStrm );
};

// ============================================================================

/** Stores all colors of the color palette. */
class ColorPalette : public GlobalDataHelper
{
public:
    /** Constructs the color palette with predefined color values. */
    explicit            ColorPalette( const GlobalDataHelper& rGlobalData );

    /** Returns true, if the passed element represents a supported palette context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Appends a user-defined color to the palette. */
    void                appendColor( sal_Int32 nRGBValue );

    /** Imports the PALETTE record from the passed stream. */
    void                importPalette( BiffInputStream& rStrm );

    /** Rturns the RGB value of the color with the passed index. */
    sal_Int32           getColor( sal_Int32 nColorId ) const;

private:
    ::std::vector< sal_Int32 > maColors;    /// List of RGB values.
    size_t              mnAppendIndex;      /// Index to append a new color.
    sal_Int32           mnWindowColor;      /// System window background color.
    sal_Int32           mnWinTextColor;     /// System window text color.
};

// ============================================================================

/** Contains all XML font attributes, e.g. from a font or rPr element. */
struct OoxFontData
{
    ::rtl::OUString     maName;             /// Font name.
    OoxColor            maColor;            /// Font color.
    sal_Int32           mnFamily;           /// Font family.
    sal_Int32           mnCharSet;          /// Windows font character set.
    double              mfHeight;           /// Font height in points.
    sal_Int32           mnUnderline;        /// Underline style.
    sal_Int32           mnEscapement;       /// Escapement style.
    bool                mbBold;             /// True = bold characters.
    bool                mbItalic;           /// True = italic characters.
    bool                mbStrikeout;        /// True = Strike out characters.
    bool                mbOutline;          /// True = outlined characters.
    bool                mbShadow;           /// True = shadowed chgaracters.

    explicit            OoxFontData();
};

// ============================================================================

class Font : public GlobalDataHelper
{
public:
    explicit            Font( const GlobalDataHelper& rGlobalData );
    explicit            Font( const GlobalDataHelper& rGlobalData, const OoxFontData& rFontData );

    /** Returns true, if the passed element represents a supported font context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Sets font formatting attributes for the passed element. */
    void                importAttribs( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );

    /** Imports the FONT record from the passed stream. */
    void                importFont( BiffInputStream& rStrm );
    /** Imports the FONTCOLOR record from the passed stream. */
    void                importFontColor( BiffInputStream& rStrm );

    /** Returns the OOX font data structure. This function can be called before
        finalizeImport() has been called. */
    inline const OoxFontData& getFontData() const { return maOoxData; }
    /** Returns the text encoding for strings used with this font. This
        function can be called before finalizeImport() has been called. */
    rtl_TextEncoding    getFontEncoding() const;

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Returns an API font descriptor with own font information. */
    const ::com::sun::star::awt::FontDescriptor& getFontDescriptor() const;
    /** Returns true, if the font requires rich text formatting in Calc.
        @descr  Example: Font escapement is a cell attribute in Excel, but Calc
        needs an rich text cell for this attribute. */
    bool                needsRichTextFormat() const;

    /** Writes all font attributes to the passed property set. */
    void                writeToPropertySet(
                            ::oox::core::PropertySet& rPropSet,
                            FontPropertyType ePropType ) const;

private:
    /** Reads and sets height and flags. */
    void                importFontData2( BiffInputStream& rStrm );
    /** Reads and sets weight, escapement, underline, family, charset (BIFF5+). */
    void                importFontData5( BiffInputStream& rStrm );

    /** Reads and sets a byte string as font name. */
    void                importFontName2( BiffInputStream& rStrm );
    /** Reads and sets a Unicode string as font name. */
    void                importFontName8( BiffInputStream& rStrm );

private:
    OoxFontData         maOoxData;
    ApiFontData         maApiData;
};

typedef ::boost::shared_ptr< Font > FontRef;

// ============================================================================

/** Contains all XML cell alignment attributes, e.g. from an alignment element. */
struct OoxAlignmentData
{
    sal_Int32           mnHorAlign;         /// Horizontal alignment.
    sal_Int32           mnVerAlign;         /// Vertical alignment.
    sal_Int32           mnTextDir;          /// CTL text direction.
    sal_Int32           mnRotation;         /// Text rotation angle.
    sal_Int32           mnIndent;           /// Indentation.
    bool                mbWrapText;         /// True = multi-line text.
    bool                mbShrink;           /// True = shrink to fit cell size.

    explicit            OoxAlignmentData();

    /** Sets horizontal alignment from the passed BIFF data. */
    void                setBiffHorAlign( sal_uInt8 nHorAlign );
    /** Sets vertical alignment from the passed BIFF data. */
    void                setBiffVerAlign( sal_uInt8 nVerAlign );
    /** Sets rotation from the passed BIFF text orientation. */
    void                setBiffTextOrient( sal_uInt8 nTextOrient );
};

// ============================================================================

class Alignment : public GlobalDataHelper
{
public:
    explicit            Alignment( const GlobalDataHelper& rGlobalData );

    /** Sets all attributes from the alignment element. */
    void                importAlignment( const ::oox::core::AttributeList& rAttribs );

    /** Sets the alignment attributes from the passed BIFF2 XF record data. */
    void                setBiff2Data( sal_uInt8 nFlags );
    /** Sets the alignment attributes from the passed BIFF3 XF record data. */
    void                setBiff3Data( sal_uInt16 nAlign );
    /** Sets the alignment attributes from the passed BIFF4 XF record data. */
    void                setBiff4Data( sal_uInt16 nAlign );
    /** Sets the alignment attributes from the passed BIFF5 XF record data. */
    void                setBiff5Data( sal_uInt16 nAlign );
    /** Sets the alignment attributes from the passed BIFF8 XF record data. */
    void                setBiff8Data( sal_uInt16 nAlign, sal_uInt16 nMiscAttrib );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Returns the OOX alignment data struct. */
    inline const OoxAlignmentData& getOoxData() const { return maOoxData; }
    /** Returns the converted API alignment data struct. */
    inline const ApiAlignmentData& getApiData() const { return maApiData; }

private:
    OoxAlignmentData    maOoxData;          /// Data from alignment element.
    ApiAlignmentData    maApiData;          /// Alignment data converted to API constants.
};

// ============================================================================

/** Contains all XML cell protection attributes, e.g. from a protection element. */
struct OoxProtectionData
{
    bool                mbLocked;           /// True = locked against editing.
    bool                mbHidden;           /// True = formula is hidden.

    explicit            OoxProtectionData();
};

// ============================================================================

class Protection : public GlobalDataHelper
{
public:
    explicit            Protection( const GlobalDataHelper& rGlobalData );

    /** Sets all attributes from the protection element. */
    void                importProtection( const ::oox::core::AttributeList& rAttribs );

    /** Sets the protection attributes from the passed BIFF2 XF record data. */
    void                setBiff2Data( sal_uInt8 nNumFmt );
    /** Sets the protection attributes from the passed BIFF3-BIFF8 XF record data. */
    void                setBiff3Data( sal_uInt16 nProt );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Returns the OOX protection data struct. */
    inline const OoxProtectionData& getOoxData() const { return maOoxData; }
    /** Returns the converted API protection data struct. */
    inline const ApiProtectionData& getApiData() const { return maApiData; }

private:
    OoxProtectionData   maOoxData;          /// Data from protection element.
    ApiProtectionData   maApiData;          /// Protection data converted to API constants.
};

// ============================================================================

/** Contains XML attributes of a single border line. */
struct OoxBorderLineData
{
    OoxColor            maColor;            /// Borderline color.
    sal_Int32           mnStyle;            /// Border line style.
    bool                mbUsed;             /// True = line format used.

    explicit            OoxBorderLineData();

    /** Sets line style and line color from the passed BIFF data. */
    void                setBiffData( sal_uInt8 nLineStyle, sal_uInt16 nLineColor );
};

// ----------------------------------------------------------------------------

/** Contains XML attributes of a complete cell border. */
struct OoxBorderData
{
    OoxBorderLineData   maLeft;             /// Left line format.
    OoxBorderLineData   maRight;            /// Right line format.
    OoxBorderLineData   maTop;              /// Top line format.
    OoxBorderLineData   maBottom;           /// Bottom line format.
    OoxBorderLineData   maDiagonal;         /// Diagonal line format.
    bool                mbDiagTLtoBR;       /// True = top-left to bottom-right on.
    bool                mbDiagBLtoTR;       /// True = bottom-left to top-right on.

    explicit            OoxBorderData();
};

// ============================================================================

class Border : public GlobalDataHelper
{
public:
    explicit            Border( const GlobalDataHelper& rGlobalData );

    /** Returns true, if the passed element represents a supported border context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Sets global border attributes from the border element. */
    void                importBorder( const ::oox::core::AttributeList& rAttribs );
    /** Sets border attributes for the border line with the passed element identifier. */
    void                importStyle( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );
    /** Sets color attributes for the border line with the passed element identifier. */
    void                importColor( sal_Int32 nElement, const ::oox::core::AttributeList& rAttribs );

    /** Sets the border attributes from the passed BIFF2 XF record data. */
    void                setBiff2Data( sal_uInt8 nFlags );
    /** Sets the border attributes from the passed BIFF3/BIFF4 XF record data. */
    void                setBiff3Data( sal_uInt32 nBorder );
    /** Sets the border attributes from the passed BIFF5 XF record data. */
    void                setBiff5Data( sal_uInt32 nBorder, sal_uInt32 nArea );
    /** Sets the border attributes from the passed BIFF8 XF record data. */
    void                setBiff8Data( sal_uInt32 nBorder1, sal_uInt32 nBorder2 );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Writes all border attributes to the passed property set. */
    void                writeToPropertySet( ::oox::core::PropertySet& rPropSet ) const;

private:
    OoxBorderLineData*  getBorderLine( sal_Int32 nElement );

    /** Converts border line data to an API struct, returns true, if the line is marked as used. */
    bool                convertBorderLine(
                            ::com::sun::star::table::BorderLine& rBorderLine,
                            const OoxBorderLineData& rLineData );

private:
    OoxBorderData       maOoxData;
    ApiBorderData       maApiData;
};

typedef ::boost::shared_ptr< Border > BorderRef;

// ============================================================================

/** Contains XML pattern fill attributes from the patternFill element. */
struct OoxPatternFillData
{
    OoxColor            maPatternColor;     /// Pattern foreground color.
    OoxColor            maFillColor;        /// Background fill color.
    sal_Int32           mnPattern;          /// Pattern identifier (e.g. solid).
    bool                mbPatternUsed;      /// True = pattern used.
    bool                mbPattColorUsed;    /// True = pattern foreground color used.
    bool                mbFillColorUsed;    /// True = background fill color used.

    explicit            OoxPatternFillData( bool bUsedFlagsDefault );

    /** Sets the pattern and pattern colors from the passed BIFF data. */
    void                setBiffData( sal_uInt16 nPatternColor, sal_uInt16 nFillColor, sal_uInt8 nPattern );
};

// ----------------------------------------------------------------------------

/** Contains XML gradient fill attributes from the gradientFill element. */
struct OoxGradientFillData
{
    typedef ::oox::core::RefMap< sal_Int32, OoxColor > OoxColorMap;

    OoxColorMap         maColors;           /// Gradient colors.

    explicit            OoxGradientFillData();
};

// ============================================================================

/** Contains cell fill attributes, either a pattern fill or a gradient fill. */
class Fill : public GlobalDataHelper
{
public:
    explicit            Fill( const GlobalDataHelper& rGlobalData );

    /** Returns true, if the passed element represents a supported fill context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Sets attributes of a patternFill element. */
    void                importPatternFill( const ::oox::core::AttributeList& rAttribs, bool bUsedFlagsDefault );
    /** Sets the pattern color from the fgColor element. */
    void                importFgColor( const ::oox::core::AttributeList& rAttribs );
    /** Sets the background color from the bgColor element. */
    void                importBgColor( const ::oox::core::AttributeList& rAttribs );
    /** Sets attributes of a gradientFill element. */
    void                importGradientFill( const ::oox::core::AttributeList& rAttribs );
    /** Sets a color from the color element in a gradient fill. */
    void                importColor( const ::oox::core::AttributeList& rAttribs, sal_Int32 nPosition );

    /** Sets the fill attributes from the passed BIFF2 XF record data. */
    void                setBiff2Data( sal_uInt8 nFlags );
    /** Sets the fill attributes from the passed BIFF3/BIFF4 XF record data. */
    void                setBiff3Data( sal_uInt16 nArea );
    /** Sets the fill attributes from the passed BIFF5 XF record data. */
    void                setBiff5Data( sal_uInt32 nArea );
    /** Sets the fill attributes from the passed BIFF8 XF record data. */
    void                setBiff8Data( sal_uInt32 nBorder2, sal_uInt16 nArea );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Writes all fill attributes to the passed property set. */
    void                writeToPropertySet( ::oox::core::PropertySet& rPropSet ) const;

private:
    typedef ::boost::shared_ptr< OoxPatternFillData >   OoxPatternRef;
    typedef ::boost::shared_ptr< OoxGradientFillData >  OoxGradientRef;

    OoxPatternRef       mxOoxPattData;
    OoxGradientRef      mxOoxGradData;
    ApiSolidFillData    maApiData;
};

typedef ::boost::shared_ptr< Fill > FillRef;

// ============================================================================

/** Contains all data for a cell format or cell style. */
struct OoxXfData
{
    sal_Int32           mnStyleXfId;        /// Index to parent style XF.
    sal_Int32           mnFontId;           /// Index to font data list.
    sal_Int32           mnNumFmtId;         /// Index to number format list.
    sal_Int32           mnBorderId;         /// Index to list of cell borders.
    sal_Int32           mnFillId;           /// Index to list of cell areas.
    bool                mbCellXf;           /// True = cell XF, false = style XF.
    bool                mbAlignUsed;        /// True = alignment used.
    bool                mbProtUsed;         /// True = cell protection used.
    bool                mbFontUsed;         /// True = font index used.
    bool                mbNumFmtUsed;       /// True = number format used.
    bool                mbBorderUsed;       /// True = border data used.
    bool                mbAreaUsed;         /// True = area data used.

    explicit            OoxXfData();
};

// ============================================================================

/** Represents a cell format or a cell style (called XF, extended format).

    This class stores the type (cell/style), the index to the parent style (if
    it is a cell format) and all "attribute used" flags, which reflect the
    state of specific attribute groups (true = user has changed the attributes)
    and all formatting data.
 */
class Xf : public GlobalDataHelper
{
public:
    explicit            Xf( const GlobalDataHelper& rGlobalData );

    /** Sets all "attribute used" flags to the passed state. */
    void                setAllUsedFlags( bool bUsed );

    /** Returns true, if the passed element represents a supported XF context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Sets all attributes from the cellXf or cellStyleXf element. */
    void                importCellXf( const ::oox::core::AttributeList& rAttribs, bool bCellXf );
    /** Sets all attributes from the alignment element. */
    void                importAlignment( const ::oox::core::AttributeList& rAttribs );
    /** Sets all attributes from the protection element. */
    void                importProtection( const ::oox::core::AttributeList& rAttribs );

    /** Imports the XF record from the passed stream. */
    void                importXf( BiffInputStream& rStrm );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Returns the referred font object. */
    FontRef             getFont() const;
    /** Returns the alignment data of this style. */
    inline const Alignment& getAlignment() const { return maAlignment; }
    /** Returns the cell protection data of this style. */
    inline const Protection& getProtection() const { return maProtection; }
    /** Returns true, if any "attribute used" flags are ste in this XF. */
    bool                hasAnyUsedFlags() const;

    /** Writes all formatting attributes to the passed property set. */
    void                writeToPropertySet( ::oox::core::PropertySet& rPropSet ) const;

private:
    /** Sets "attribute used" flags from the passed BIFF bit field. */
    void                setUsedFlags( sal_uInt8 nUsedFlags );
    /** Updates own used flags from the passed cell style XF. */
    void                updateUsedFlags( const Xf& rStyleXf );

private:
    OoxXfData           maOoxData;          /// Data from cellXf or cellStyleXf element.
    Alignment           maAlignment;        /// Cell alignment data.
    Protection          maProtection;       /// Cell protection data.
};

typedef ::boost::shared_ptr< Xf > XfRef;

// ============================================================================

class Dxf : public GlobalDataHelper
{
public:
    explicit            Dxf( const GlobalDataHelper& rGlobalData );

    void                importDxf( const ::oox::core::AttributeList& rAttribs );
    FontRef             getFont();
    FillRef             getFill();
    BorderRef           getBorder();

private:
    FontRef             mpFont;
    FillRef             mpFill;
    BorderRef           mpBorder;
};

typedef ::boost::shared_ptr< Dxf > DxfRef;

// ============================================================================

/** Contains attributes of a cell style, e.g. from the cellStyle element. */
struct OoxCellStyleData
{
    ::rtl::OUString     maName;             /// Cell style name.
    sal_Int32           mnBuiltinId;        /// Identifier for builtin styles.
    sal_Int32           mnLevel;            /// Level for builtin column/row styles.
    bool                mbCustom;           /// True = customized builtin style.
    bool                mbHidden;           /// True = style not visible in GUI.

    explicit            OoxCellStyleData();

    /** Returns true, if this style is a builtin style. */
    inline bool         isBuiltin() const { return mnBuiltinId >= 0; }
    /** Returns true, if this style represents the default document cell style. */
    inline bool         isDefaultStyle() const { return mnBuiltinId == OOX_STYLE_NORMAL; }
    /** Returns the style name used in the UI. */
    ::rtl::OUString     createStyleName() const;
};

// ============================================================================

class CellStyle : public GlobalDataHelper
{
public:
    explicit            CellStyle( const GlobalDataHelper& rGlobalData );

    /** Returns true, if the passed element represents a supported palette context. */
    static bool         isSupportedContext( sal_Int32 nElement, sal_Int32 nParentContext );
    /** Imports passed attributes from the cellStyle element. */
    void                importCellStyle( const ::oox::core::AttributeList& rAttribs );
    /** Imports style settings from a STYLE BIFF record. */
    void                importStyle( BiffInputStream& rStrm, bool bBuiltin );

    /** Returns true, if this style represents the default document cell style. */
    bool                isDefaultStyle() const;

    /** Creates the style sheet described by the style XF with the passed identifier. */
    const ::rtl::OUString& createStyleSheet( sal_Int32 nXfId, bool bSkipDefaultBuiltin = false );

private:
    OoxCellStyleData    maOoxData;
    ::rtl::OUString     maFinalName;        /// Final style name used in API.
};

typedef ::boost::shared_ptr< CellStyle > CellStyleRef;

// ============================================================================

class StylesBuffer : public GlobalDataHelper
{
public:
    explicit            StylesBuffer( const GlobalDataHelper& rGlobalData );

    /** Creates a new empty border object.
        @param opnBorderId  (out-param) The identifier of the new border object. */
    BorderRef           createBorder( sal_Int32* opnBorderId = 0 );
    /** Creates a new empty fill object.
        @param opnFillId  (out-param) The identifier of the new fill object. */
    FillRef             createFill( sal_Int32* opnFillId = 0 );

    /** Appends a new color to the color palette. */
    void                importPaletteColor( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new font object. */
    FontRef             importFont( const ::oox::core::AttributeList& rAttribs );
    /** Inserts a new number format code. */
    void                importNumFmt( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new cell border object. */
    BorderRef           importBorder( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new cell fill object. */
    FillRef             importFill( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new cell format or cell style object. */
    XfRef               importXf( sal_Int32 nContext, const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new differential cell format object. */
    DxfRef              importDxf( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a new named cell style object. */
    CellStyleRef        importCellStyle( const ::oox::core::AttributeList& rAttribs );

    /** Imports the PALETTE record from the passed stream. */
    void                importPalette( BiffInputStream& rStrm );
    /** Imports the FONT record from the passed stream. */
    void                importFont( BiffInputStream& rStrm );
    /** Imports the FONTCOLOR record from the passed stream. */
    void                importFontColor( BiffInputStream& rStrm );
    /** Imports the FORMAT record from the passed stream. */
    void                importFormat( BiffInputStream& rStrm );
    /** Imports the XF record from the passed stream. */
    void                importXf( BiffInputStream& rStrm );
    /** Imports the STYLE record from the passed stream. */
    void                importStyle( BiffInputStream& rStrm );

    /** Final processing after import of all style settings. */
    void                finalizeImport();

    /** Returns the RGB value of the passed color object. */
    sal_Int32           getColor( const OoxColor& rColor ) const;
    /** Returns the specified font object. */
    FontRef             getFont( sal_Int32 nFontId ) const;
    /** Returns the specified cell format object. */
    XfRef               getCellXf( sal_Int32 nXfId ) const;
    /** Returns the specified style format object. */
    XfRef               getStyleXf( sal_Int32 nXfId ) const;
    /** Returns the specified diferential cell format object. */
    DxfRef              getDxf( sal_Int32 nDxfId ) const;

    /** Returns the font object of the specified cell XF. */
    FontRef             getFontFromCellXf( sal_Int32 nXfId ) const;
    /** Returns the default application font (used in the "Normal" cell style). */
    FontRef             getDefaultFont() const;
    /** Returns the default application font data (used in the "Normal" cell style). */
    const OoxFontData&  getDefaultFontData() const;

    /** Creates the style sheet described by the style XF with the passed identifier. */
    const ::rtl::OUString& createStyleSheet( sal_Int32 nXfId ) const;

    /** Writes the font attributes of the specified font data to the passed property set. */
    void                writeFontToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nFontId ) const;
    /** Writes the specified number format to the passed property set. */
    void                writeNumFmtToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nNumFmtId ) const;
    /** Writes the border attributes of the specified border data to the passed property set. */
    void                writeBorderToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nBorderId ) const;
    /** Writes the fill attributes of the specified fill data to the passed property set. */
    void                writeFillToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nFillId ) const;
    /** Writes the cell formatting attributes of the specified XF to the passed property set. */
    void                writeCellXfToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nXfId ) const;
    /** Writes the cell formatting attributes of the specified style XF to the passed property set. */
    void                writeStyleXfToPropertySet( ::oox::core::PropertySet& rPropSet, sal_Int32 nXfId ) const;

private:
    typedef ::oox::core::RefVector< Font >              FontVec;
    typedef ::oox::core::RefVector< Border >            BorderVec;
    typedef ::oox::core::RefVector< Fill >              FillVec;
    typedef ::oox::core::RefVector< Xf >                XfVec;
    typedef ::oox::core::RefVector< Dxf >               DxfVec;
    typedef ::oox::core::RefMap< sal_Int32, CellStyle > CellStyleMap;

    ColorPalette        maPalette;          /// Color palette.
    FontVec             maFonts;            /// List of font objects.
    NumberFormatsBuffer maNumFmts;          /// List of all number format codes.
    BorderVec           maBorders;          /// List of cell border objects.
    FillVec             maFills;            /// List of cell area fill objects.
    XfVec               maCellXfs;          /// List of cell formats.
    XfVec               maStyleXfs;         /// List of cell styles.
    DxfVec              maDxfs;             /// List of differential cell styles.
    CellStyleMap        maCellStyles;       /// List of named cell styles.
    ::rtl::OUString     maDefStyleName;     /// API name of default cell style.
    sal_Int32           mnDefStyleXf;       /// Style XF index of default cell style.
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

