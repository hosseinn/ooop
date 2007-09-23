/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: globaldatahelper.hxx,v $
 *
 *  $Revision: 1.1.2.32 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 12:29:40 $
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

#ifndef OOX_XLS_GLOBALDATAHELPER_HXX
#define OOX_XLS_GLOBALDATAHELPER_HXX

#include <boost/shared_ptr.hpp>
#include "oox/core/xmlfilterbase.hxx"
#include "oox/core/binaryfilterbase.hxx"
#include "oox/xls/biffhelper.hxx"

namespace com { namespace sun { namespace star {
    namespace awt { class XDevice; }
    namespace table { struct CellAddress; }
    namespace sheet { class XSpreadsheetDocument; }
    namespace sheet { class XSpreadsheet; }
    namespace sheet { class XNamedRanges; }
} } }

namespace oox {
namespace xls {

// ============================================================================

/** An enumeration for all supported spreadsheet filter types. */
enum FilterType
{
    FILTER_OOX,         /// MS Excel OOX (Office Open XML).
    FILTER_BIFF,        /// MS Excel BIFF (Binary Interchange File Format).
    FILTER_UNKNOWN      /// Unknown filter type.
};

// ============================================================================

#if OSL_DEBUG_LEVEL > 0

class GlobalDataHelperDebug
{
public:
    inline explicit     GlobalDataHelperDebug( sal_Int32& rnCount ) : mrnCount( rnCount ) { ++mrnCount; }
    inline explicit     GlobalDataHelperDebug( const GlobalDataHelperDebug& rCopy ) : mrnCount( rCopy.mrnCount ) { ++mrnCount; }
    virtual             ~GlobalDataHelperDebug() { --mrnCount; }
private:
    sal_Int32&          mrnCount;
};

#endif

// ============================================================================

struct GlobalData;
typedef ::boost::shared_ptr< GlobalData > GlobalDataRef;

class ViewSettings;
class WorksheetBuffer;
class ThemeBuffer;
class StylesBuffer;
class SharedStringsBuffer;
class CondFormatBuffer;
class ExternalLinkBuffer;
class DefinedNamesBuffer;
class WebQueryBuffer;
class PivotTableBuffer;
class FormulaParser;
class UnitConverter;
class AddressConverter;
class StylesPropertyHelper;
class PageStylePropertyHelper;

/** Helper class to provice access to global workbook data.

    All classes derived from this helper class will have access to a singleton
    object (a GlobalData struct) containing global workbook data, e.g. strings
    and cell formatting.
 */
class GlobalDataHelper
#if OSL_DEBUG_LEVEL > 0
    : private GlobalDataHelperDebug
#endif
{
public:
    /** Initial construction of all global objects. */
    explicit            GlobalDataHelper( GlobalData& rGlobalData );

    virtual             ~GlobalDataHelper();

    /** Creates a global data object to be used in an OOX filter. */
    static GlobalDataRef createGlobalDataStruct( const ::oox::core::XmlFilterRef& rxFilter );

    /** Creates a global data object to be used in a BIFF filter. */
    static GlobalDataRef createGlobalDataStruct( const ::oox::core::BinaryFilterRef& rxFilter, BiffType eBiff );

    // global data and document model -----------------------------------------

    /** Return this global data helper for better code readability in derived classes. */
    inline const GlobalDataHelper& getGlobalData() const { return *this; }

    /** Returns the base filter object (base class of all filters). */
    const ::oox::core::FilterBase& getBaseFilter() const;

    /** Returns the file type of the current filter. */
    FilterType          getFilterType() const;

    /** Returns true, if the file is a multi-sheet document, or false if single-sheet. */
    bool                isWorkbookFile() const;

    /** Returns a reference to the source/target spreadsheet document model. */
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheetDocument >
                        getDocument() const;

    /** Returns a reference to the specified spreadsheet in the document model. */
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XSpreadsheet >
                        getSheet( sal_Int32 nSheet ) const;

    /** Returns the reference device of the document. */
    ::com::sun::star::uno::Reference< ::com::sun::star::awt::XDevice >
                        getReferenceDevice() const;

    /** Returns the container for defined names from the Calc document. */
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XNamedRanges >
                        getNamedRanges() const;

    // buffers ----------------------------------------------------------------

    /** Returns the workbook and sheet view settings object. */
    ViewSettings&       getViewSettings() const;

    /** Returns the worksheet buffer containing sheet names and properties. */
    WorksheetBuffer&    getWorksheets() const;

    /** Returns the office theme object read from the theme substorage. */
    ThemeBuffer&        getTheme() const;

    /** Returns all cell formatting objects read from the styles substream. */
    StylesBuffer&       getStyles() const;

    /** Returns the shared strings read from the shared strings substream. */
    SharedStringsBuffer& getSharedStrings() const;

    /** Returns the conditional formatting items. */
    CondFormatBuffer&   getCondFormats() const;

    /** Returns the external links read from the external links substream. */
    ExternalLinkBuffer& getExternalLinks() const;

    /** Returns the defined names read from the workbook globals. */
    DefinedNamesBuffer& getDefinedNames() const;

    /** Returns the web queries. */
    WebQueryBuffer&     getWebQueries() const;

    /** Returns the pivot tables. */
    PivotTableBuffer&   getPivotTables() const;

    // converters -------------------------------------------------------------

    /** Returns the import formula parser. */
    FormulaParser&      getFormulaParser() const;

    /** Returns the measurement unit converter. */
    UnitConverter&      getUnitConverter() const;

    /** Returns the converter for string to cell address/range conversion. */
    AddressConverter&   getAddressConverter() const;

    /** Returns the converter for all styles related properties. */
    StylesPropertyHelper& getStylesPropertyHelper() const;

    /** Returns the converter for all page style related properties. */
    PageStylePropertyHelper& getPageStylePropertyHelper() const;

    // OOX specific -----------------------------------------------------------

    /** Returns the base OOX filter object.
        Must not be called, if current filter is not the OOX filter. */
    const ::oox::core::XmlFilterRef& getOoxFilter() const;

    // BIFF specific ----------------------------------------------------------

    /** Returns the base BIFF filter object. */
    const ::oox::core::BinaryFilterRef& getBiffFilter() const;
    /** Returns the BIFF type in binary filter. */
    BiffType            getBiff() const;

    /** Returns the text encoding used to import/export byte strings. */
    rtl_TextEncoding    getTextEncoding() const;
    /** Sets the text encoding to import/export byte strings. */
    void                setTextEncoding( rtl_TextEncoding eTextEnc );
    /** Sets code page read from a CODEPAGE record for byte string import. */
    void                setCodePage( sal_uInt16 nCodePage );
    /** Sets text encoding from the default application font, if CODEPAGE record is missing. */
    void                setAppFontEncoding( rtl_TextEncoding eAppFontEnc );

    /** Enables workbook file mode, used for BIFF4 workspace files. */
    void                setIsWorkbookFile();
    /** Recreates global buffers that are used per sheet in specific BIFF versions. */
    void                createBuffersPerSheet();

    /** Looks for a password provided via API, or queries it via GUI. */
    ::rtl::OUString     queryPassword() const;

private:
    GlobalData&         mrGlobalData;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

