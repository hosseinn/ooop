/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: globaldatahelper.cxx,v $
 *
 *  $Revision: 1.1.2.37 $
 *
 *  last change: $Author: dr $ $Date: 2007/09/05 12:31:23 $
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

#include "oox/xls/globaldatahelper.hxx"
#include <osl/thread.h>
#include <osl/time.h>
#include <rtl/strbuf.hxx>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/awt/XDevice.hpp>
#include <com/sun/star/document/XActionLockable.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/sheet/XNamedRanges.hpp>
#include "oox/core/propertyset.hxx"
#include "oox/xls/addressconverter.hxx"
#include "oox/xls/condformatbuffer.hxx"
#include "oox/xls/defnamesbuffer.hxx"
#include "oox/xls/externallinkbuffer.hxx"
#include "oox/xls/formulaparser.hxx"
#include "oox/xls/pagestyle.hxx"
#include "oox/xls/pivottablebuffer.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"
#include "oox/xls/stylesbuffer.hxx"
#include "oox/xls/stylespropertyhelper.hxx"
#include "oox/xls/themebuffer.hxx"
#include "oox/xls/unitconverter.hxx"
#include "oox/xls/viewsettings.hxx"
#include "oox/xls/webquerybuffer.hxx"
#include "oox/xls/worksheetbuffer.hxx"

using ::rtl::OUString;
using ::rtl::OStringBuffer;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::Exception;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::uno::UNO_QUERY_THROW;
using ::com::sun::star::container::XIndexAccess;
using ::com::sun::star::awt::XDevice;
using ::com::sun::star::document::XActionLockable;
using ::com::sun::star::table::CellAddress;
using ::com::sun::star::sheet::XSpreadsheetDocument;
using ::com::sun::star::sheet::XSpreadsheet;
using ::com::sun::star::sheet::XNamedRanges;
using ::oox::core::FilterBase;
using ::oox::core::XmlFilterRef;
using ::oox::core::BinaryFilterRef;
using ::oox::core::PropertySet;

// Set this define to 1 to show the load/save time of a document in an assertion.
#define OOX_SHOW_LOADSAVE_TIME 0

namespace oox {
namespace xls {

// ============================================================================

#if OSL_DEBUG_LEVEL > 0

struct GlobalDataDebug
{
#if OOX_SHOW_LOADSAVE_TIME
    TimeValue           maStartTime;
#endif
    sal_Int32           mnDebugCount;

    explicit            GlobalDataDebug();
                        ~GlobalDataDebug();
};

GlobalDataDebug::GlobalDataDebug() :
    mnDebugCount( 0 )
{
#if OOX_SHOW_LOADSAVE_TIME
    osl_getSystemTime( &maStartTime );
#endif
}

GlobalDataDebug::~GlobalDataDebug()
{
#if OOX_SHOW_LOADSAVE_TIME
    TimeValue aEndTime;
    osl_getSystemTime( &aEndTime );
    sal_Int32 nMillis = (aEndTime.Seconds - maStartTime.Seconds) * 1000 + static_cast< sal_Int32 >( aEndTime.Nanosec - maStartTime.Nanosec ) / 1000000;
    OSL_ENSURE( false, OStringBuffer( "load/save time = " ).append( nMillis / 1000.0 ).append( " seconds" ).getStr() );
#endif
    OSL_ENSURE( mnDebugCount == 0, "GlobalDataDebug::~GlobalDataDebug - failed to delete some objects" );
}

#endif

// ============================================================================

/** Contains global workbook data, e.g. strings and cell formatting. */
struct GlobalData
#if OSL_DEBUG_LEVEL > 0
    : public GlobalDataDebug
#endif
{
    typedef ::std::auto_ptr< ViewSettings >             ViewSettingsPtr;
    typedef ::std::auto_ptr< WorksheetBuffer >          WorksheetBfrPtr;
    typedef ::std::auto_ptr< ThemeBuffer >              ThemeBfrPtr;
    typedef ::std::auto_ptr< StylesBuffer >             StylesBfrPtr;
    typedef ::std::auto_ptr< SharedStringsBuffer >      SharedStrBfrPtr;
    typedef ::std::auto_ptr< CondFormatBuffer >         CondFormatBfrPtr;
    typedef ::std::auto_ptr< ExternalLinkBuffer >       ExtLinkBfrPtr;
    typedef ::std::auto_ptr< DefinedNamesBuffer >       DefNamesBfrPtr;
    typedef ::std::auto_ptr< WebQueryBuffer >           WebQueryBfrPtr;
    typedef ::std::auto_ptr< PivotTableBuffer >         PivotTableBfrPtr;
    typedef ::std::auto_ptr< UnitConverter >            UnitConvPtr;
    typedef ::std::auto_ptr< AddressConverter >         AddressConvPtr;
    typedef ::std::auto_ptr< StylesPropertyHelper >     StylesPropHlpPtr;
    typedef ::std::auto_ptr< PageStylePropertyHelper >  PageStylePropHlpPtr;
    typedef ::std::auto_ptr< FormulaParser >            FormulaParserPtr;

    OUString            maRefDeviceProp;        /// ReferenceDevice property name.
    OUString            maNamedRangesProp;      /// NamedRanges property name.
    Reference< XSpreadsheetDocument > mxDoc;    /// Document model.
    FilterBase*         mpBaseFilter;           /// Base filter object.
    FilterType          meFilterType;           /// File type of the filter.
    bool                mbWorkbook;             /// True = multi-sheet file.

    // buffers
    ViewSettingsPtr     mxViewSettings;         /// Workbook and sheet view settings.
    WorksheetBfrPtr     mxWorksheets;           /// Sheet info buffer.
    ThemeBfrPtr         mxTheme;                /// Formatting theme from theme substream.
    StylesBfrPtr        mxStyles;               /// All cell style objects from styles substream.
    SharedStrBfrPtr     mxSharedStrings;        /// All strings from shared strings substream.
    CondFormatBfrPtr    mxCondFormats;          /// All conditional formatting items.
    ExtLinkBfrPtr       mxExtLinks;             /// All external links.
    DefNamesBfrPtr      mxDefNames;             /// All defined names.
    WebQueryBfrPtr      mxWebQueries;           /// Web queries buffer.
    PivotTableBfrPtr    mxPivotTables;          /// Pivot tables buffer.

    // converters/helpers
    FormulaParserPtr    mxFmlaParser;           /// Import formula parser.
    UnitConvPtr         mxUnitConverter;        /// General unit converter.
    AddressConvPtr      mxAddrConverter;        /// Cell address and cell range address converter.
    StylesPropHlpPtr    mxStylesPropHlp;        /// Helper for all styles properties.
    PageStylePropHlpPtr mxPageStylePropHlp;     /// Helper for page style properties.

    // OOX specific
    XmlFilterRef        mxOoxFilter;            /// Base OOX filter object.

    // BIFF specific
    BinaryFilterRef     mxBiffFilter;           /// Base BIFF filter object.
    ::rtl::OUString     maPassword;             /// Password for stream encoder/decoder.
    BiffType            meBiff;                 /// BIFF version for BIFF import/export.
    rtl_TextEncoding    meTextEnc;              /// BIFF byte string text encoding.
    bool                mbHasCodePage;          /// True = CODEPAGE record exists in imported stream.
    bool                mbHasPassword;          /// True = password already querried.

    explicit            GlobalData();
                        ~GlobalData();

    /** Initializes global data for an OOX filter. */
    bool                initOoxFilter( const XmlFilterRef& rxFilter );
    /** Initializes global data for a BIFF filter. */
    bool                initBiffFilter( const BinaryFilterRef& rxFilter, BiffType eBiff );

    /** Returns the reference device of the document. */
    Reference< XDevice > getReferenceDevice() const;
    /** Returns the container for defined names from the Calc document. */
    Reference< XNamedRanges > getNamedRanges() const;

private:
    /** Initializes some basic members and sets needed document properties. */
    bool                initialize();
    /** Finalizes the filter process (sets some needed document properties). */
    void                finalize();
};

// ----------------------------------------------------------------------------

GlobalData::GlobalData() :
    maRefDeviceProp( CREATE_OUSTRING( "ReferenceDevice" ) ),
    maNamedRangesProp( CREATE_OUSTRING( "NamedRanges" ) ),
    mpBaseFilter( 0 ),
    meFilterType( FILTER_UNKNOWN ),
    mbWorkbook( false ),
    meBiff( BIFF_UNKNOWN ),
    meTextEnc( osl_getThreadTextEncoding() ),
    mbHasCodePage( false ),
    mbHasPassword( false )
{
}

GlobalData::~GlobalData()
{
    finalize();
}

bool GlobalData::initOoxFilter( const XmlFilterRef& rxFilter )
{
    meFilterType = FILTER_OOX;
    mxOoxFilter = rxFilter;
    mpBaseFilter = rxFilter.get();
    mbWorkbook = true;
    return initialize();
}

bool GlobalData::initBiffFilter( const BinaryFilterRef& rxFilter, BiffType eBiff )
{
    meFilterType = FILTER_BIFF;
    mxBiffFilter = rxFilter;
    mpBaseFilter = rxFilter.get();
    meBiff = eBiff;
    mbWorkbook = eBiff >= BIFF5;
    return initialize();
}

Reference< XDevice > GlobalData::getReferenceDevice() const
{
    PropertySet aPropSet( mxDoc );
    Reference< XDevice > xDevice;
    aPropSet.getProperty( xDevice, maRefDeviceProp );
    return xDevice;
}

Reference< XNamedRanges > GlobalData::getNamedRanges() const
{
    PropertySet aPropSet( mxDoc );
    Reference< XNamedRanges > xNamedRanges;
    aPropSet.getProperty( xNamedRanges, maNamedRangesProp );
    return xNamedRanges;
}

bool GlobalData::initialize()
{
    OSL_ENSURE( mpBaseFilter, "GlobalData::initialize - no base filter" );
    mxDoc.set( mpBaseFilter->getModel(), UNO_QUERY );
    OSL_ENSURE( mxDoc.is(), "GlobalData::initialize - no spreadsheet document" );

    // set some document properties needed during import
    if( mpBaseFilter->isImportFilter() )
    {
        PropertySet aPropSet( mxDoc );
        // #i76026# disable Undo while loading the document
        aPropSet.setProperty( CREATE_OUSTRING( "IsUndoEnabled" ), false );
        // #i79826# disable calculating automatic row height while loading the document
        aPropSet.setProperty( CREATE_OUSTRING( "IsAdjustHeightEnabled" ), false );
        // #i79890# disable automatic update of defined names
        Reference< XActionLockable > xLockable( getNamedRanges(), UNO_QUERY );
        if( xLockable.is() )
            xLockable->addActionLock();
    }

    return mxDoc.is();
}

void GlobalData::finalize()
{
    // set some document properties needed after import
    if( mpBaseFilter->isImportFilter() )
    {
        PropertySet aPropSet( mxDoc );
        // #i74668# do not insert default sheets
        aPropSet.setProperty( CREATE_OUSTRING( "IsLoaded" ), true );
        // #i79890# enable automatic update of defined names (before IsAdjustHeightEnabled!)
        Reference< XActionLockable > xLockable( getNamedRanges(), UNO_QUERY );
        if( xLockable.is() )
            xLockable->removeActionLock();
        // #i79826# enable updating automatic row height after loading the document
        aPropSet.setProperty( CREATE_OUSTRING( "IsAdjustHeightEnabled" ), true );
        // #i76026# enable Undo after loading the document
        aPropSet.setProperty( CREATE_OUSTRING( "IsUndoEnabled" ), true );
    }
}

// ----------------------------------------------------------------------------

GlobalDataHelper::GlobalDataHelper( GlobalData& rGlobalData ) :
#if OSL_DEBUG_LEVEL > 0
    GlobalDataHelperDebug( rGlobalData.mnDebugCount ),
#endif
    mrGlobalData( rGlobalData )
{
    /*  These objects cannot be constructed in the GlobalData constructor,
        because they need a living GlobalDataHelper object. */
    mrGlobalData.mxViewSettings.reset( new ViewSettings( getGlobalData() ) );
    mrGlobalData.mxWorksheets.reset( new WorksheetBuffer( getGlobalData() ) );
    mrGlobalData.mxTheme.reset( new ThemeBuffer( getGlobalData() ) );
    mrGlobalData.mxStyles.reset( new StylesBuffer( getGlobalData() ) );
    mrGlobalData.mxSharedStrings.reset( new SharedStringsBuffer( getGlobalData() ) );
    mrGlobalData.mxCondFormats.reset( new CondFormatBuffer( getGlobalData() ) );
    mrGlobalData.mxDefNames.reset( new DefinedNamesBuffer( getGlobalData() ) );
    mrGlobalData.mxExtLinks.reset( new ExternalLinkBuffer( getGlobalData() ) );
    mrGlobalData.mxWebQueries.reset( new WebQueryBuffer( getGlobalData() ) );
    mrGlobalData.mxPivotTables.reset( new PivotTableBuffer( getGlobalData() ) );

    mrGlobalData.mxUnitConverter.reset( new UnitConverter( getGlobalData() ) );
    mrGlobalData.mxAddrConverter.reset( new AddressConverter( getGlobalData() ) );
    mrGlobalData.mxStylesPropHlp.reset( new StylesPropertyHelper( getGlobalData() ) );
    mrGlobalData.mxPageStylePropHlp.reset( new PageStylePropertyHelper( getGlobalData() ) );

    if( getBaseFilter().isImportFilter() )
    {
        // import filter specific
        mrGlobalData.mxFmlaParser.reset( new FormulaParser( getGlobalData() ) );
    }
    else if( getBaseFilter().isExportFilter() )
    {
        // export filter specific
    }
}

GlobalDataHelper::~GlobalDataHelper()
{
}

GlobalDataRef GlobalDataHelper::createGlobalDataStruct( const XmlFilterRef& rxFilter )
{
    GlobalDataRef xGlobalData( new GlobalData );
    if( !xGlobalData->initOoxFilter( rxFilter ) )
        xGlobalData.reset();
    return xGlobalData;
}

GlobalDataRef GlobalDataHelper::createGlobalDataStruct( const BinaryFilterRef& rxFilter, BiffType eBiff )
{
    GlobalDataRef xGlobalData( new GlobalData );
    if( !xGlobalData->initBiffFilter( rxFilter, eBiff ) )
        xGlobalData.reset();
    return xGlobalData;
}

// global data and document model ---------------------------------------------

const FilterBase& GlobalDataHelper::getBaseFilter() const
{
    return *mrGlobalData.mpBaseFilter;
}

FilterType GlobalDataHelper::getFilterType() const
{
    return mrGlobalData.meFilterType;
}

bool GlobalDataHelper::isWorkbookFile() const
{
    return mrGlobalData.mbWorkbook;
}

Reference< XSpreadsheetDocument > GlobalDataHelper::getDocument() const
{
    return mrGlobalData.mxDoc;
}

Reference< XSpreadsheet > GlobalDataHelper::getSheet( sal_Int32 nSheet ) const
{
    Reference< XSpreadsheet > xSheet;
    try
    {
        Reference< XIndexAccess > xSheetsIA( getDocument()->getSheets(), UNO_QUERY_THROW );
        xSheet.set( xSheetsIA->getByIndex( nSheet ), UNO_QUERY_THROW );
    }
    catch( Exception& )
    {
    }
    return xSheet;
}

Reference< XDevice > GlobalDataHelper::getReferenceDevice() const
{
    return mrGlobalData.getReferenceDevice();
}

Reference< XNamedRanges > GlobalDataHelper::getNamedRanges() const
{
    return mrGlobalData.getNamedRanges();
}

// buffers --------------------------------------------------------------------

ViewSettings& GlobalDataHelper::getViewSettings() const
{
    return *mrGlobalData.mxViewSettings;
}

WorksheetBuffer& GlobalDataHelper::getWorksheets() const
{
    return *mrGlobalData.mxWorksheets;
}

ThemeBuffer& GlobalDataHelper::getTheme() const
{
    return *mrGlobalData.mxTheme;
}

StylesBuffer& GlobalDataHelper::getStyles() const
{
    return *mrGlobalData.mxStyles;
}

SharedStringsBuffer& GlobalDataHelper::getSharedStrings() const
{
    return *mrGlobalData.mxSharedStrings;
}

CondFormatBuffer& GlobalDataHelper::getCondFormats() const
{
    return *mrGlobalData.mxCondFormats;
}

ExternalLinkBuffer& GlobalDataHelper::getExternalLinks() const
{
    return *mrGlobalData.mxExtLinks;
}

DefinedNamesBuffer& GlobalDataHelper::getDefinedNames() const
{
    return *mrGlobalData.mxDefNames;
}

WebQueryBuffer& GlobalDataHelper::getWebQueries() const
{
    return *mrGlobalData.mxWebQueries;
}

PivotTableBuffer& GlobalDataHelper::getPivotTables() const
{
    return *mrGlobalData.mxPivotTables;
}

// converters -----------------------------------------------------------------

FormulaParser& GlobalDataHelper::getFormulaParser() const
{
    return *mrGlobalData.mxFmlaParser;
}

UnitConverter& GlobalDataHelper::getUnitConverter() const
{
    return *mrGlobalData.mxUnitConverter;
}

AddressConverter& GlobalDataHelper::getAddressConverter() const
{
    return *mrGlobalData.mxAddrConverter;
}

StylesPropertyHelper& GlobalDataHelper::getStylesPropertyHelper() const
{
    return *mrGlobalData.mxStylesPropHlp;
}

PageStylePropertyHelper& GlobalDataHelper::getPageStylePropertyHelper() const
{
    return *mrGlobalData.mxPageStylePropHlp;
}

// OOX specific ---------------------------------------------------------------

const XmlFilterRef& GlobalDataHelper::getOoxFilter() const
{
    OSL_ENSURE( mrGlobalData.mxOoxFilter.is(), "GlobalDataHelper::getOoxFilter - missing filter" );
    return mrGlobalData.mxOoxFilter;
}

// BIFF specific --------------------------------------------------------------

const BinaryFilterRef& GlobalDataHelper::getBiffFilter() const
{
    OSL_ENSURE( mrGlobalData.mxBiffFilter.is(), "GlobalDataHelper::getBiffFilter - missing filter" );
    return mrGlobalData.mxBiffFilter;
}

BiffType GlobalDataHelper::getBiff() const
{
    return mrGlobalData.meBiff;
}

rtl_TextEncoding GlobalDataHelper::getTextEncoding() const
{
    return mrGlobalData.meTextEnc;
}

void GlobalDataHelper::setTextEncoding( rtl_TextEncoding eTextEnc )
{
    if( eTextEnc != RTL_TEXTENCODING_DONTKNOW )
        mrGlobalData.meTextEnc = eTextEnc;
}

void GlobalDataHelper::setCodePage( sal_uInt16 nCodePage )
{
    setTextEncoding( BiffHelper::calcTextEncodingFromCodePage( nCodePage ) );
    mrGlobalData.mbHasCodePage = true;
}

void GlobalDataHelper::setAppFontEncoding( rtl_TextEncoding eAppFontEnc )
{
    if( !mrGlobalData.mbHasCodePage )
        setTextEncoding( eAppFontEnc );
}

void GlobalDataHelper::setIsWorkbookFile()
{
    OSL_ENSURE( mrGlobalData.meBiff == BIFF4, "GlobalDataHelper::setIsWorkbookFile - invalid call" );
    mrGlobalData.mbWorkbook = true;
}

void GlobalDataHelper::createBuffersPerSheet()
{
    switch( mrGlobalData.meBiff )
    {
        case BIFF2:
        case BIFF3:
        break;

        case BIFF4:
            // #i11183# sheets in BIFF4W files have own styles or names
            if( isWorkbookFile() )
            {
                mrGlobalData.mxStyles.reset( new StylesBuffer( getGlobalData() ) );
                mrGlobalData.mxDefNames.reset( new DefinedNamesBuffer( getGlobalData() ) );
                mrGlobalData.mxExtLinks.reset( new ExternalLinkBuffer( getGlobalData() ) );
            }
        break;

        case BIFF5:
            // BIFF5 stores external references per sheet
            mrGlobalData.mxExtLinks.reset( new ExternalLinkBuffer( getGlobalData() ) );
        break;

        case BIFF8:
        break;

        case BIFF_UNKNOWN:
        break;
    }
}

OUString GlobalDataHelper::queryPassword() const
{
    if( !mrGlobalData.mbHasPassword )
    {
        //! TODO
        mrGlobalData.maPassword = OUString();
        // set to true, even if dialog has been cancelled (never ask twice)
        mrGlobalData.mbHasPassword = true;
    }
    return mrGlobalData.maPassword;
}

// ============================================================================

} // namespace xls
} // namespace oox

