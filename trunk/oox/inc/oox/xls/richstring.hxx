/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: richstring.hxx,v $
 *
 *  $Revision: 1.1.2.3 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/31 11:40:43 $
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

#ifndef OOX_XLS_RICHSTRING_HXX
#define OOX_XLS_RICHSTRING_HXX

#include "oox/core/containerhelper.hxx"
#include "oox/xls/stylesbuffer.hxx"

namespace com { namespace sun { namespace star {
    namespace text { class XText; }
} } }

namespace oox {
namespace xls {

// ============================================================================

const sal_uInt16 BIFF_PHONETIC_TYPE_KATAKANA_HALF   = 0;
const sal_uInt16 BIFF_PHONETIC_TYPE_KATAKANA_FULL   = 1;
const sal_uInt16 BIFF_PHONETIC_TYPE_HIRAGANA        = 2;

const sal_uInt16 BIFF_PHONETIC_ALIGN_NONE           = 0;
const sal_uInt16 BIFF_PHONETIC_ALIGN_LEFT           = 1;
const sal_uInt16 BIFF_PHONETIC_ALIGN_CENTER         = 2;
const sal_uInt16 BIFF_PHONETIC_ALIGN_DISTRIBUTED    = 3;

// ============================================================================

/** Contains text data and font attributes for a part of a rich formatted string. */
class RichStringPortion : public GlobalDataHelper
{
public:
    explicit            RichStringPortion( const GlobalDataHelper& rGlobalData );

    /** Sets text data for this portion. */
    void                setText( const ::rtl::OUString& rText );
    /** Creates and returns a new font formatting object. */
    FontRef             importFont( const ::oox::core::AttributeList& rAttribs );
    /** Links this portion to a font object from the global font list. */
    void                setFontId( sal_Int32 nFontId );

    /** Final processing after import of all strings. */
    void                finalizeImport();

    /** Converts the portion and appends it to the passed XText. */
    void                convert(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::text::XText >& rxText,
                            sal_Int32 nXfId );

private:
    ::rtl::OUString     maText;         /// Portion text.
    FontRef             mxFont;         /// Embedded portion font, may be empty.
    sal_Int32           mnFontId;       /// Link to global font list.
};

typedef ::boost::shared_ptr< RichStringPortion > RichStringPortionRef;

// ============================================================================

/** Contains text data and positioning information for a phonetic text portion. */
class RichStringPhonetic : public GlobalDataHelper
{
public:
    explicit            RichStringPhonetic( const GlobalDataHelper& rGlobalData );

    /** Sets text data for this phonetic portion. */
    void                setText( const ::rtl::OUString& rText );
    /** Imports attributes of a phonetic run (rPh element). */
    void                importPhoneticRun( const ::oox::core::AttributeList& rAttribs );
    /** Sets the associated range in base text for this phonetic portion. */
    void                setBaseRange( sal_Int32 nBasePos, sal_Int32 nBaseEnd );

private:
    ::rtl::OUString     maText;         /// Portion text.
    sal_Int32           mnBasePos;      /// Start position in base text.
    sal_Int32           mnBaseEnd;      /// One-past-end position in base text.
};

typedef ::boost::shared_ptr< RichStringPhonetic > RichStringPhoneticRef;

// ============================================================================

class PhoneticSettings : public GlobalDataHelper
{
public:
    explicit            PhoneticSettings( const GlobalDataHelper& rGlobalData );

    /** Imports phonetic settings from the rPhoneticPr element. */
    void                importPhoneticPr( const ::oox::core::AttributeList& rAttribs );
    /** Imports phonetic settings from the passed BIFF Unicode string data. */
    void                setBiffData( sal_uInt16 nFlags, sal_uInt16 nFontId );

private:
    sal_Int32           mnType;         /// Phonetic text type.
    sal_Int32           mnAlignment;    /// Phonetic portion alignment.
    sal_Int32           mnFontId;       /// Font identifier for text formatting.
};

// ============================================================================

/** Represents a position in a rich-string containing current font identifier.

    This object stores the position of a formatted character in a rich-string
    and the identifier of a font from the global font list used to format this
    and the following characters.
 */
struct BiffRichStringFontId
{
    sal_Int32           mnPos;          /// First character in the string.
    sal_Int32           mnFontId;       /// Font identifier for the next characters.

    explicit inline     BiffRichStringFontId() : mnPos( 0 ), mnFontId( -1 ) {}
    explicit inline     BiffRichStringFontId( sal_Int32 nPos, sal_Int32 nFontId ) :
                            mnPos( nPos ), mnFontId( nFontId ) {}
};

/** A vector with all font identifiers in a rich-string. */
typedef ::std::vector< BiffRichStringFontId > BiffRichStringFontIdVec;

// ============================================================================

/** Represents a phonetic text portion in a rich-string with phonetic text. */
struct BiffPhoneticPortion
{
    sal_Int32           mnPos;          /// First character in phonetic text.
    sal_Int32           mnBasePos;      /// First character in base text.
    sal_Int32           mnBaseLen;      /// Number of characters in base text.

    explicit inline     BiffPhoneticPortion() : mnPos( -1 ), mnBasePos( -1 ), mnBaseLen( 0 ) {}
    explicit inline     BiffPhoneticPortion( sal_Int32 nPos, sal_Int32 nBasePos, sal_Int32 nBaseLen ) :
                            mnPos( nPos ), mnBasePos( nBasePos ), mnBaseLen( nBaseLen ) {}
};

/** A vector with all phonetic portions in a rich-string. */
typedef ::std::vector< BiffPhoneticPortion > BiffPhoneticPortionVec;

// ============================================================================

/** Contains string data and a list of formatting runs for a rich formatted string. */
class RichString : public GlobalDataHelper
{
public:
    explicit            RichString( const GlobalDataHelper& rGlobalData );

    /** Appends and returns a portion object for a plain string (t element). */
    RichStringPortionRef importText( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a portion object for a new formatting run (r element). */
    RichStringPortionRef importRun( const ::oox::core::AttributeList& rAttribs );
    /** Appends and returns a phonetic text object for a new phonetic run (rPh element). */
    RichStringPhoneticRef importPhoneticRun( const ::oox::core::AttributeList& rAttribs );
    /** Imports phonetic settings from the rPhoneticPr element. */
    void                importPhoneticPr( const ::oox::core::AttributeList& rAttribs );

    /** Appends a BIFF rich-string font identifier to the passed vector. */
    static void         appendFontId( BiffRichStringFontIdVec& orFontIds, sal_Int32 nPos, sal_Int32 nFontId );
    /** Reads nCount font identifiers from stream and inserts them into the passed vector. */
    static void         importFontIds( BiffRichStringFontIdVec& orFontIds, BiffInputStream& rStrm, sal_uInt16 nCount, bool b16Bit );
    /** Reads count and font identifiers from stream and inserts them into the passed vector. */
    static void         importFontIds( BiffRichStringFontIdVec& orFontIds, BiffInputStream& rStrm, bool b16Bit );

    /** Imports a byte string from the passed BIFF stream. */
    void                importByteString( BiffInputStream& rStrm, rtl_TextEncoding eDefaultTextEnc, BiffStringFlags nFlags = BIFF_STR_DEFAULT );
    /** Imports a Unicode rich-string from the passed BIFF stream. */
    void                importUniString( BiffInputStream& rStrm, BiffStringFlags nFlags = BIFF_STR_DEFAULT );

    /** Final processing after import of all strings. */
    void                finalizeImport();

    /** Converts the string and writes it into the passed XText. */
    void                convert(
                            const ::com::sun::star::uno::Reference< ::com::sun::star::text::XText >& rxText,
                            sal_Int32 nXfId ) const;

private:
    /** Creates, appends, and returns a new empty string portion. */
    RichStringPortionRef createPortion();
    /** Creates, appends, and returns a new empty phonetic text portion. */
    RichStringPhoneticRef createPhonetic();

    /** Appends a BIFF rich-string phonetic portion to the passed vector. */
    static void         appendPhoneticPortion( BiffPhoneticPortionVec& orPhonetics, sal_Int32 nPos, sal_Int32 nBasePos, sal_Int32 nBaseLen );
    /** Reads nCount phonetic portions from stream and inserts them into the passed vector. */
    static void         importPhoneticPortions( BiffPhoneticPortionVec& orPhonetics, BiffInputStream& rStrm, sal_uInt16 nCount );
    /** Reads phonetic data from the passed stream. */
    ::rtl::OUString     importPhonetic( BiffPhoneticPortionVec& orPhonetics, BiffInputStream& rStrm, sal_uInt32 nPhoneticSize );

    /** Create base text portions from the passed string and character formatting. */
    void                createBiffPortions( const ::rtl::OString& rText, rtl_TextEncoding eDefaultTextEnc, BiffRichStringFontIdVec& rFontIds );
    /** Create base text portions from the passed string and character formatting. */
    void                createBiffPortions( const ::rtl::OUString& rText, BiffRichStringFontIdVec& rFontIds );
    /** Create phonetic text portions from the passed string and portion data. */
    void                createBiffPhonetics( const ::rtl::OUString& rText, BiffPhoneticPortionVec& rPhonetics, sal_Int32 nBaseLen );

private:
    typedef ::oox::core::RefVector< RichStringPortion >     PortionVec;
    typedef ::oox::core::RefVector< RichStringPhonetic >    PhoneticVec;

    PortionVec          maPortions;     /// String portions with font data.
    PhoneticVec         maPhonetics;    /// Phonetic text portions.
    PhoneticSettings    maPhoneticSett; /// Phonetic settings for this string.
};

typedef ::boost::shared_ptr< RichString > RichStringRef;

// ============================================================================

} // namespace xls
} // namespace oox

#endif

