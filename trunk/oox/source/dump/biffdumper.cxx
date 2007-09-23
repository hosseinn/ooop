/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffdumper.cxx,v $
 *
 *  $Revision: 1.1.2.36 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/24 09:08:49 $
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

#include "oox/dump/biffdumper.hxx"

#include <osl/thread.h>
#include <rtl/tencinfo.h>
#include <rtl/strbuf.hxx>
#include <com/sun/star/util/DateTime.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include "oox/core/binaryinputstream.hxx"
#include "oox/xls/biffinputstream.hxx"
#include "oox/xls/biffformatdetector.hxx"
#include "oox/xls/formulabase.hxx"
#include "oox/xls/sheetdatacontext.hxx"
#include "oox/xls/sheetviewscontext.hxx"
#include "oox/xls/workbookfragment.hxx"
#include "oox/xls/worksheetfragment.hxx"

#if OOX_INCLUDE_DUMPER

using ::rtl::OUString;
using ::rtl::OUStringBuffer;
using ::rtl::OString;
using ::rtl::OStringBuffer;
using ::com::sun::star::uno::Reference;
using ::com::sun::star::uno::UNO_QUERY;
using ::com::sun::star::util::DateTime;
using ::com::sun::star::sheet::XSpreadsheetDocument;
using ::oox::core::BinaryInputStream;
using ::oox::core::FilterBase;

using namespace ::oox::xls;

namespace oox {
namespace dump {
namespace biff {

namespace  {

// special dumper characters --------------------------------------------------

const sal_Unicode OOX_DUMP_BASECLASS        = 'B';
const sal_Unicode OOX_DUMP_FUNCSEP          = ',';
const sal_Unicode OOX_DUMP_LISTSEP          = ',';
const sal_Unicode OOX_DUMP_ADDRABS          = '$';
const sal_Unicode OOX_DUMP_R1C1ROW          = 'R';
const sal_Unicode OOX_DUMP_R1C1COL          = 'C';
const sal_Unicode OOX_DUMP_R1C1OPEN         = '[';
const sal_Unicode OOX_DUMP_R1C1CLOSE        = ']';
const sal_Unicode OOX_DUMP_RANGESEP         = ':';
const sal_Unicode OOX_DUMP_TABSEP           = '!';
const sal_Unicode OOX_DUMP_ARRAYSEP         = ';';
const sal_Unicode OOX_DUMP_EMPTYVALUE       = '~';
const sal_Unicode OOX_DUMP_CMDPROMPT        = '?';
const sal_Unicode OOX_DUMP_PLACEHOLDER      = '\x01';

// record identifiers unsupported in filter -----------------------------------

const sal_uInt16 BIFF2_ID_DIMENSIONS        = 0x0000;
const sal_uInt16 BIFF3_ID_DIMENSIONS        = 0x0200;

const sal_uInt16 BIFF_ID_NOTE               = 0x001C;
const sal_uInt16 BIFF_ID_OBJ                = 0x005D;
const sal_uInt16 BIFF_ID_MSODRAWINGGROUP    = 0x00EB;
const sal_uInt16 BIFF_ID_MSODRAWING         = 0x00EC;
const sal_uInt16 BIFF_ID_MSODRAWINGSEL      = 0x00ED;
const sal_uInt16 BIFF_ID_TXO                = 0x01B6;

const sal_uInt16 BIFF_ID_CHCHART            = 0x1002;
const sal_uInt16 BIFF_ID_CHSERIES           = 0x1003;
const sal_uInt16 BIFF_ID_CHDATAFORMAT       = 0x1006;
const sal_uInt16 BIFF_ID_CHLINEFORMAT       = 0x1007;
const sal_uInt16 BIFF_ID_CHMARKERFORMAT     = 0x1009;
const sal_uInt16 BIFF_ID_CHAREAFORMAT       = 0x100A;
const sal_uInt16 BIFF_ID_CHPIEFORMAT        = 0x100B;
const sal_uInt16 BIFF_ID_CHATTACHEDLABEL    = 0x100C;
const sal_uInt16 BIFF_ID_CHSTRING           = 0x100D;
const sal_uInt16 BIFF_ID_CHTYPEGROUP        = 0x1014;
const sal_uInt16 BIFF_ID_CHLEGEND           = 0x1015;
const sal_uInt16 BIFF_ID_CHBAR              = 0x1017;
const sal_uInt16 BIFF_ID_CHLINE             = 0x1018;
const sal_uInt16 BIFF_ID_CHAREA             = 0x101A;
const sal_uInt16 BIFF_ID_CHPIE              = 0x1019;
const sal_uInt16 BIFF_ID_CHSCATTER          = 0x101B;
const sal_uInt16 BIFF_ID_CHCHARTLINE        = 0x101C;
const sal_uInt16 BIFF_ID_CHAXIS             = 0x101D;
const sal_uInt16 BIFF_ID_CHTICK             = 0x101E;
const sal_uInt16 BIFF_ID_CHVALUERANGE       = 0x101F;
const sal_uInt16 BIFF_ID_CHLABELRANGE       = 0x1020;
const sal_uInt16 BIFF_ID_CHAXISLINE         = 0x1021;
const sal_uInt16 BIFF_ID_CHDEFAULTTEXT      = 0x1024;
const sal_uInt16 BIFF_ID_CHTEXT             = 0x1025;
const sal_uInt16 BIFF_ID_CHFONT             = 0x1026;
const sal_uInt16 BIFF_ID_CHOBJECTLINK       = 0x1027;
const sal_uInt16 BIFF_ID_CHFRAME            = 0x1032;
const sal_uInt16 BIFF_ID_CHBEGIN            = 0x1033;
const sal_uInt16 BIFF_ID_CHEND              = 0x1034;
const sal_uInt16 BIFF_ID_CHPLOTFRAME        = 0x1035;
const sal_uInt16 BIFF_ID_CHCHART3D          = 0x103A;
const sal_uInt16 BIFF_ID_CHPICFORMAT        = 0x103C;
const sal_uInt16 BIFF_ID_CHDROPBAR          = 0x103D;
const sal_uInt16 BIFF_ID_CHRADARLINE        = 0x103E;
const sal_uInt16 BIFF_ID_CHRADARAREA        = 0x1040;
const sal_uInt16 BIFF_ID_CHSURFACE          = 0x103F;
const sal_uInt16 BIFF_ID_CHAXESSET          = 0x1041;
const sal_uInt16 BIFF_ID_CHPROPERTIES       = 0x1044;
const sal_uInt16 BIFF_ID_CHSERGROUP         = 0x1045;
const sal_uInt16 BIFF5_ID_CHPIVOTREF        = 0x1048;
const sal_uInt16 BIFF8_ID_CHPIVOTREF        = 0x0858;
const sal_uInt16 BIFF_ID_CHSERPARENT        = 0x104A;
const sal_uInt16 BIFF_ID_CHSERTRENDLINE     = 0x104B;
const sal_uInt16 BIFF_ID_CHFORMAT           = 0x104E;
const sal_uInt16 BIFF_ID_CHFRAMEPOS         = 0x104F;
const sal_uInt16 BIFF_ID_CHFORMATRUNS       = 0x1050;
const sal_uInt16 BIFF_ID_CHSOURCELINK       = 0x1051;
const sal_uInt16 BIFF_ID_CHSERERRORBAR      = 0x105B;
const sal_uInt16 BIFF_ID_CHSERIESFORMAT     = 0x105D;
const sal_uInt16 BIFF_ID_CH3DDATAFORMAT     = 0x105F;
const sal_uInt16 BIFF_ID_CHPIEEXT           = 0x1061;
const sal_uInt16 BIFF_ID_CHESCHERFORMAT     = 0x1066;
const sal_uInt16 BIFF_ID_CHWRAPPEDRECORD    = 0x0851;
const sal_uInt16 BIFF_ID_CHUNITPROPERTIES   = 0x0857;
const sal_uInt16 BIFF_ID_CHLABELRANGE2      = 0x1062;
const sal_uInt16 BIFF_ID_CHPLOTGROWTH       = 0x1064;

const sal_uInt16 BIFF_ID_DCONREF            = 0x0051;
const sal_uInt16 BIFF_ID_DCONNAME           = 0x0052;
const sal_uInt16 BIFF_ID_SXVIEW             = 0x00B0;
const sal_uInt16 BIFF_ID_SXVD               = 0x00B1;
const sal_uInt16 BIFF_ID_SXVI               = 0x00B2;
const sal_uInt16 BIFF_ID_SXIVD              = 0x00B4;
const sal_uInt16 BIFF_ID_SXLI               = 0x00B5;
const sal_uInt16 BIFF_ID_SXDI               = 0x00C5;
const sal_uInt16 BIFF_ID_SXDB               = 0x00C6;
const sal_uInt16 BIFF_ID_SXFIELD            = 0x00C7;
const sal_uInt16 BIFF_ID_SXSTRING           = 0x00CD;
const sal_uInt16 BIFF_ID_SXDATETIME         = 0x00CE;
const sal_uInt16 BIFF_ID_SXEXT              = 0x00DC;
const sal_uInt16 BIFF_ID_SXVDEX             = 0x0100;

const sal_uInt16 BIFF_ID_OBJEND             = 0x0000;   /// End of OBJ.
const sal_uInt16 BIFF_ID_OBJMACRO           = 0x0004;   /// Macro link.
const sal_uInt16 BIFF_ID_OBJGMO             = 0x0006;   /// Group marker.
const sal_uInt16 BIFF_ID_OBJCF              = 0x0007;   /// Clipboard format.
const sal_uInt16 BIFF_ID_OBJPIOGRBIT        = 0x0008;   /// Option flags.
const sal_uInt16 BIFF_ID_OBJPICTFMLA        = 0x0009;   /// OLE link formula.
const sal_uInt16 BIFF_ID_OBJCBLS            = 0x000A;   /// Check box/radio button data.
const sal_uInt16 BIFF_ID_OBJSBS             = 0x000C;   /// Scroll bar data.
const sal_uInt16 BIFF_ID_OBJSBSFMLA         = 0x000E;   /// Scroll bar/list box/combo box cell link.
const sal_uInt16 BIFF_ID_OBJGBODATA         = 0x000F;   /// Group box data.
const sal_uInt16 BIFF_ID_OBJLBSDATA         = 0x0013;   /// List box/combo box data.
const sal_uInt16 BIFF_ID_OBJCBLSFMLA        = 0x0014;   /// Check box/radio button cell link.
const sal_uInt16 BIFF_ID_OBJCMO             = 0x0015;   /// Common object settings.

// constants unsupported in filter --------------------------------------------

const sal_uInt16 BIFF_OBJPIO_MANUALSIZE     = 0x0001;
const sal_uInt16 BIFF_OBJPIO_LINKED         = 0x0002;
const sal_uInt16 BIFF_OBJPIO_SYMBOL         = 0x0008;
const sal_uInt16 BIFF_OBJPIO_CONTROL        = 0x0010;   /// Form control.
const sal_uInt16 BIFF_OBJPIO_CTLSSTREAM     = 0x0020;   /// Data in Ctls stream.
const sal_uInt16 BIFF_OBJPIO_AUTOLOAD       = 0x0200;

const sal_uInt16 BIFF_OBJCMO_GROUP          = 0x0000;
const sal_uInt16 BIFF_OBJCMO_LINE           = 0x0001;
const sal_uInt16 BIFF_OBJCMO_RECTANGLE      = 0x0002;
const sal_uInt16 BIFF_OBJCMO_ELLIPSE        = 0x0003;
const sal_uInt16 BIFF_OBJCMO_ARC            = 0x0004;
const sal_uInt16 BIFF_OBJCMO_CHART          = 0x0005;
const sal_uInt16 BIFF_OBJCMO_TEXT           = 0x0006;
const sal_uInt16 BIFF_OBJCMO_BUTTON         = 0x0007;
const sal_uInt16 BIFF_OBJCMO_PICTURE        = 0x0008;
const sal_uInt16 BIFF_OBJCMO_POLYGON        = 0x0009;
const sal_uInt16 BIFF_OBJCMO_CHECKBOX       = 0x000B;
const sal_uInt16 BIFF_OBJCMO_OPTIONBUTTON   = 0x000C;
const sal_uInt16 BIFF_OBJCMO_EDIT           = 0x000D;
const sal_uInt16 BIFF_OBJCMO_LABEL          = 0x000E;
const sal_uInt16 BIFF_OBJCMO_DIALOG         = 0x000F;
const sal_uInt16 BIFF_OBJCMO_SPIN           = 0x0010;
const sal_uInt16 BIFF_OBJCMO_SCROLLBAR      = 0x0011;
const sal_uInt16 BIFF_OBJCMO_LISTBOX        = 0x0012;
const sal_uInt16 BIFF_OBJCMO_GROUPBOX       = 0x0013;
const sal_uInt16 BIFF_OBJCMO_COMBOBOX       = 0x0014;
const sal_uInt16 BIFF_OBJCMO_NOTE           = 0x0019;
const sal_uInt16 BIFF_OBJCMO_DRAWING        = 0x001E;

const sal_uInt16 BIFF_PT_NOSTRING           = 0xFFFF;

} // namespace

// ============================================================================
// ============================================================================

namespace {

inline sal_Int32 lclReadCol( BiffInputStream& rStrm, bool bCol16Bit )
{
    return bCol16Bit ? rStrm.readuInt16() : rStrm.readuInt8();
}

inline sal_Int32 lclReadRow( BiffInputStream& rStrm, bool bRow32Bit )
{
    return bRow32Bit ? rStrm.readInt32() : rStrm.readuInt16();
}

} // namespace

// ----------------------------------------------------------------------------

void Address::read( BiffInputStream& rStrm, bool bCol16Bit, bool bRow32Bit )
{
    mnRow = lclReadRow( rStrm, bRow32Bit );
    mnCol = lclReadCol( rStrm, bCol16Bit );
}

// ----------------------------------------------------------------------------

void Range::read( BiffInputStream& rStrm, bool bCol16Bit, bool bRow32Bit )
{
    maFirst.mnRow = lclReadRow( rStrm, bRow32Bit );
    maLast.mnRow = lclReadRow( rStrm, bRow32Bit );
    maFirst.mnCol = lclReadCol( rStrm, bCol16Bit );
    maLast.mnCol = lclReadCol( rStrm, bCol16Bit );
}

// ----------------------------------------------------------------------------

void RangeList::read( BiffInputStream& rStrm, bool bCol16Bit, bool bRow32Bit )
{
    sal_uInt16 nCount;
    rStrm >> nCount;
    resize( nCount );
    for( iterator aIt = begin(), aEnd = end(); rStrm.isValid() && (aIt != aEnd); ++aIt )
        aIt->read( rStrm, bCol16Bit, bRow32Bit );
}

// ============================================================================

void StringHelper::appendAddrCol( OUStringBuffer& rStr, sal_Int32 nCol, bool bRel )
{
    if( !bRel ) rStr.append( OOX_DUMP_ADDRABS );
    sal_Int32 nPos = rStr.getLength();
    for( sal_Int32 nTemp = nCol; nTemp >= 0; (nTemp /= 26) -= 1 )
        rStr.insert( nPos, static_cast< sal_Unicode >( 'A' + (nTemp % 26) ) );
}

void StringHelper::appendAddrRow( OUStringBuffer& rStr, sal_Int32 nRow, bool bRel )
{
    if( !bRel ) rStr.append( OOX_DUMP_ADDRABS );
    appendDec( rStr, nRow + 1 );
}

void StringHelper::appendAddrName( OUStringBuffer& rStr, sal_Unicode cPrefix, sal_Int32 nColRow, bool bRel )
{
    rStr.append( cPrefix );
    if( bRel && (nColRow != 0) )
    {
        rStr.append( OOX_DUMP_R1C1OPEN );
        appendDec( rStr, nColRow );
        rStr.append( OOX_DUMP_R1C1CLOSE );
    }
    else if( !bRel )
        appendDec( rStr, nColRow + 1 );
}

void StringHelper::appendAddress( OUStringBuffer& rStr, const Address& rPos )
{
    appendAddrCol( rStr, rPos.mnCol, true );
    appendAddrRow( rStr, rPos.mnRow, true );
}

void StringHelper::appendRange( OUStringBuffer& rStr, const Range& rRange )
{
    appendAddress( rStr, rRange.maFirst );
    rStr.append( OOX_DUMP_RANGESEP );
    appendAddress( rStr, rRange.maLast );
}

void StringHelper::appendRangeList( OUStringBuffer& rStr, const RangeList& rRanges )
{
    OUStringBuffer aData;
    for( RangeList::const_iterator aIt = rRanges.begin(), aEnd = rRanges.end(); aIt != aEnd; ++aIt )
    {
        OUStringBuffer aRange;
        appendRange( aRange, *aIt );
        appendToken( aData, aRange.makeStringAndClear(), OOX_DUMP_LISTSEP );
    }
    rStr.append( aData.makeStringAndClear() );
}

void StringHelper::appendAddress( OUStringBuffer& rStr, const TokenAddress& rPos, bool bNameMode )
{
    if( bNameMode && (rPos.mbRelCol || rPos.mbRelRow) )
    {
        appendAddrName( rStr, OOX_DUMP_R1C1ROW, rPos.mnRow, rPos.mbRelRow );
        appendAddrName( rStr, OOX_DUMP_R1C1COL, rPos.mnCol, rPos.mbRelCol );
    }
    else
    {
        appendAddrCol( rStr, rPos.mnCol, rPos.mbRelCol );
        appendAddrRow( rStr, rPos.mnRow, rPos.mbRelRow );
    }
}

void StringHelper::appendRange( OUStringBuffer& rStr, const TokenRange& rRange, bool bNameMode )
{
    appendAddress( rStr, rRange.maFirst, bNameMode );
    rStr.append( OOX_DUMP_RANGESEP );
    appendAddress( rStr, rRange.maLast, bNameMode );
}

// ============================================================================
// ============================================================================

class BiffStreamInput : public Input
{
public:
    inline explicit     BiffStreamInput( BiffInputStream& rStrm ) : mrStrm( rStrm ) {}
    virtual             ~BiffStreamInput();

    virtual sal_Int64   getSize() const;
    virtual sal_Int64   tell() const;
    virtual void        seek( sal_Int64 nPos );
    virtual void        skip( sal_Int32 nBytes );
    virtual sal_Int32   read( void* pBuffer, sal_Int32 nBytes );

    virtual BiffStreamInput& operator>>( sal_Int8& rnData );
    virtual BiffStreamInput& operator>>( sal_uInt8& rnData );
    virtual BiffStreamInput& operator>>( sal_Int16& rnData );
    virtual BiffStreamInput& operator>>( sal_uInt16& rnData );
    virtual BiffStreamInput& operator>>( sal_Int32& rnData );
    virtual BiffStreamInput& operator>>( sal_uInt32& rnData );
    virtual BiffStreamInput& operator>>( float& rfData );
    virtual BiffStreamInput& operator>>( double& rfData );

private:
    BiffInputStream&    mrStrm;
};

// ----------------------------------------------------------------------------

BiffStreamInput::~BiffStreamInput()
{
}

sal_Int64 BiffStreamInput::getSize() const
{
    return mrStrm.getRecSize();
}

sal_Int64 BiffStreamInput::tell() const
{
    return mrStrm.getRecPos();
}

void BiffStreamInput::seek( sal_Int64 nPos )
{
    mrStrm.seek( static_cast< sal_uInt32 >( nPos ) );
}

void BiffStreamInput::skip( sal_Int32 nBytes )
{
    mrStrm.ignore( static_cast< sal_uInt32 >( nBytes ) );
}

sal_Int32 BiffStreamInput::read( void* pBuffer, sal_Int32 nSize )
{
    return static_cast< sal_Int32 >( mrStrm.read( pBuffer, static_cast< sal_uInt32 >( nSize ) ) );
}

BiffStreamInput& BiffStreamInput::operator>>( sal_Int8& rnData )   { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( sal_uInt8& rnData )  { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( sal_Int16& rnData )  { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( sal_uInt16& rnData ) { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( sal_Int32& rnData )  { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( sal_uInt32& rnData ) { mrStrm >> rnData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( float& rfData )      { mrStrm >> rfData; return *this; }
BiffStreamInput& BiffStreamInput::operator>>( double& rfData )     { mrStrm >> rfData; return *this; }

// ============================================================================
// ============================================================================

RecordHeaderObject::RecordHeaderObject( const InputObjectBase& rParent )
{
    static const RecordHeaderConfigInfo saHeaderCfgInfo =
    {
        "RECORD-NAMES",
        "show-record-pos",
        "show-record-size",
        "show-record-id",
        "show-record-name",
        "show-record-body",
    };
    RecordHeaderBase::construct( rParent, saHeaderCfgInfo );
    if( RecordHeaderBase::implIsValid() )
        mbMergeContRec = cfg().getBoolOption( "merge-continue-record", true );
}

void RecordHeaderObject::dumpRecordHeader( BiffInputStream& rStrm )
{
    MultiItemsGuard aMultiGuard( out() );
    writeEmptyItem( "REC" );
    if( isShowRecPos() )  writeHexItem( "pos", static_cast< sal_uInt32 >( rStrm.getCoreStreamPos() - 4 ) );
    if( isShowRecSize() ) writeHexItem( "size", rStrm.getRecSize() );
    if( isShowRecId() )   writeHexItem( "id", rStrm.getRecId() );
    if( isShowRecName() ) writeNameItem( "name", rStrm.getRecId(), getRecNames() );
}

bool RecordHeaderObject::implIsValid() const
{
    return RecordHeaderBase::implIsValid();
}

// ============================================================================
// ============================================================================

BiffConfig::BiffConfig( const Config& rParent, BiffType eBiff ) :
    meBiff( eBiff )
{
    Config::construct( rParent );
}

bool BiffConfig::implIsValid() const
{
    return (meBiff != BIFF_UNKNOWN) && Config::implIsValid();
}

NameListRef BiffConfig::implGetNameList( const OUString& rKey ) const
{
    NameListRef xList = Config::implGetNameList( rKey );
    if( !xList )
    {
        OUString aBaseKey = rKey + CREATE_OUSTRING( "-BIFF" );
        switch( meBiff )
        {
            // fall-through intended!
            case BIFF8: if( !xList ) xList = Config::implGetNameList( aBaseKey + OUString( sal_Unicode( '8' ) ) );
            case BIFF5: if( !xList ) xList = Config::implGetNameList( aBaseKey + OUString( sal_Unicode( '5' ) ) );
            case BIFF4: if( !xList ) xList = Config::implGetNameList( aBaseKey + OUString( sal_Unicode( '4' ) ) );
            case BIFF3: if( !xList ) xList = Config::implGetNameList( aBaseKey + OUString( sal_Unicode( '3' ) ) );
            case BIFF2: if( !xList ) xList = Config::implGetNameList( aBaseKey + OUString( sal_Unicode( '2' ) ) );
            case BIFF_UNKNOWN: break;
        }
    }
    return xList;
}

// ============================================================================

BiffSharedData::BiffSharedData( BiffType eBiff ) :
    meBiff( eBiff ),
    meTextEnc( osl_getThreadTextEncoding() )
{
}

BiffSharedData::~BiffSharedData()
{
}

void BiffSharedData::setTextEncoding( rtl_TextEncoding eTextEnc )
{
    if( eTextEnc != RTL_TEXTENCODING_DONTKNOW )
        meTextEnc = eTextEnc;
}

sal_uInt16 BiffSharedData::getFontCount() const
{
    return static_cast< sal_uInt16 >( maFontEncs.size() );
}

rtl_TextEncoding BiffSharedData::getFontEncoding( sal_uInt16 nFontId ) const
{
    return (nFontId < getFontCount()) ? maFontEncs[ nFontId ] : meTextEnc;
}

void BiffSharedData::appendFontEncoding( rtl_TextEncoding eFontEnc )
{
    maFontEncs.push_back( (eFontEnc == RTL_TEXTENCODING_DONTKNOW) ? meTextEnc : eFontEnc );
    if( maFontEncs.size() == 4 )
        maFontEncs.push_back( meTextEnc );
}

sal_uInt16 BiffSharedData::getXfCount() const
{
    return static_cast< sal_uInt16 >( maXfFontIds.size() );
}

rtl_TextEncoding BiffSharedData::getXfEncoding( sal_uInt16 nXfId ) const
{
    sal_uInt16 nFontId = (nXfId < getXfCount()) ? maXfFontIds[ nXfId ] : 0;
    return getFontEncoding( nFontId );
}

void BiffSharedData::appendXfFontId( sal_uInt16 nFontId )
{
    maXfFontIds.push_back( nFontId );
}

bool BiffSharedData::implIsValid() const
{
    return meBiff != BIFF_UNKNOWN;
}

// ============================================================================

BiffObjectBase::BiffObjectBase()
{
}

BiffObjectBase::~BiffObjectBase()
{
}

void BiffObjectBase::construct( const ObjectBase& rParent, BinaryInputStreamRef xStrm, BiffType eBiff )
{
    StreamObjectRef xStrmObj( new BinaryStreamObject( rParent, xStrm ) );
    WrappedStreamObject::construct( rParent, xStrmObj );
    constructBiffObjBase( eBiff );
}

void BiffObjectBase::construct( const OleStorageObject& rParentStrg, const OUString& rStrmName, BiffType eBiff )
{
    StreamObjectRef xStrmObj( new OleStreamObject( rParentStrg, rStrmName ) );
    WrappedStreamObject::construct( rParentStrg, xStrmObj );
    constructBiffObjBase( eBiff );
}

void BiffObjectBase::construct( const BiffObjectBase& rParent )
{
    *this = rParent;
}

bool BiffObjectBase::implIsValid() const
{
    return isValid( mxBiffCfg ) && isValid( mxBiffData ) && mxStrm.get() && WrappedStreamObject::implIsValid();
}

void BiffObjectBase::implDumpExtendedHeader()
{
    WrappedStreamObject::implDumpExtendedHeader();
    writeNameItem( "biff-version", static_cast< sal_uInt32 >( getBiff() ), "BIFF" );
}

ConfigRef BiffObjectBase::implReconstructConfig()
{
    mxBiffCfg.reset( new BiffConfig( cfg(), getBiff() ) );
    return mxBiffCfg;
}

InputRef BiffObjectBase::implReconstructInput()
{
    InputRef xNewIn;
    if( mxStrm.get() )
        xNewIn.reset( new BiffStreamInput( *mxStrm ) );
    return xNewIn;
}

OUString BiffObjectBase::getErrorName( sal_uInt8 nErrCode ) const
{
    return cfg().getName( mxErrCodes, nErrCode );
}

// ----------------------------------------------------------------------------

double BiffObjectBase::writeRkItem( const sal_Char* pcName, sal_Int32 nRk )
{
    MultiItemsGuard aMultiGuard( out() );
    writeHexItem( pcName, static_cast< sal_uInt32 >( nRk ), "RK-FLAGS" );
    double fValue = BiffHelper::calcDoubleFromRk( nRk );
    writeDecItem( "decoded", fValue );
    return fValue;
}

void BiffObjectBase::writeBooleanItem( const sal_Char* pcName, sal_uInt8 nBool )
{
    writeDecItem( pcName, nBool, mxBoolean );
}

void BiffObjectBase::writeErrorCodeItem( const sal_Char* pcName, sal_uInt8 nErrCode )
{
    writeHexItem( pcName, nErrCode, mxErrCodes );
}

void BiffObjectBase::writeColIndexItem( const sal_Char* pcName, sal_Int32 nCol )
{
    Output& rOut = out();
    ItemGuard aItem( rOut, pcName );
    rOut.writeDec( nCol );
    aItem.cont();
    StringHelper::appendAddrCol( rOut.getLine(), nCol, true );
}

void BiffObjectBase::writeColRangeItem( const sal_Char* pcName, sal_Int32 nCol1, sal_Int32 nCol2 )
{
    Output& rOut = out();
    ItemGuard aItem( rOut, pcName );
    rOut.writeDec( nCol1 );
    rOut.writeChar( OOX_DUMP_RANGESEP );
    rOut.writeDec( nCol2 );
    aItem.cont();
    StringHelper::appendAddrCol( rOut.getLine(), nCol1, true );
    rOut.writeChar( OOX_DUMP_RANGESEP );
    StringHelper::appendAddrCol( rOut.getLine(), nCol2, true );
}

void BiffObjectBase::writeRowIndexItem( const sal_Char* pcName, sal_Int32 nRow )
{
    Output& rOut = out();
    ItemGuard aItem( rOut, pcName );
    rOut.writeDec( nRow );
    aItem.cont();
    StringHelper::appendAddrRow( rOut.getLine(), nRow, true );
}

void BiffObjectBase::writeRowRangeItem( const sal_Char* pcName, sal_Int32 nRow1, sal_Int32 nRow2 )
{
    Output& rOut = out();
    ItemGuard aItem( rOut, pcName );
    rOut.writeDec( nRow1 );
    rOut.writeChar( OOX_DUMP_RANGESEP );
    rOut.writeDec( nRow2 );
    aItem.cont();
    StringHelper::appendAddrRow( rOut.getLine(), nRow1, true );
    rOut.writeChar( OOX_DUMP_RANGESEP );
    StringHelper::appendAddrRow( rOut.getLine(), nRow2, true );
}

void BiffObjectBase::writeAddressItem( const sal_Char* pcName, const Address& rPos )
{
    ItemGuard aItem( out(), pcName );
    StringHelper::appendAddress( out().getLine(), rPos );
}

void BiffObjectBase::writeRangeItem( const sal_Char* pcName, const Range& rRange )
{
    ItemGuard aItem( out(), pcName );
    StringHelper::appendRange( out().getLine(), rRange );
}

void BiffObjectBase::writeRangeListItem( const sal_Char* pcName, const RangeList& rRanges )
{
    MultiItemsGuard aMultiGuard( out() );
    writeEmptyItem( pcName );
    writeDecItem( "count", static_cast< sal_uInt16 >( rRanges.size() ) );
    ItemGuard aItem( out(), "ranges" );
    StringHelper::appendRangeList( out().getLine(), rRanges );
}

void BiffObjectBase::writeRichStringFontIds( const BiffRichStringFontIdVec& rFontIds )
{
    if( !rFontIds.empty() )
    {
        writeDecItem( "font-count", static_cast< sal_uInt32 >( rFontIds.size() ) );
        TableGuard aTabGuard( out(), 14 );
        for( BiffRichStringFontIdVec::const_iterator aIt = rFontIds.begin(), aEnd = rFontIds.end(); aIt != aEnd; ++aIt )
        {
            MultiItemsGuard aMultiGuard( out() );
            writeDecItem( "char-idx", aIt->mnPos );
            writeDecItem( "font-idx", aIt->mnFontId, "FONTNAMES" );
        }
    }
}

// ----------------------------------------------------------------------------

OUString BiffObjectBase::dumpByteString( const sal_Char* pcName, BiffStringFlags nFlags, rtl_TextEncoding eDefaultTextEnc )
{
    OSL_ENSURE( !getFlag( nFlags, static_cast< BiffStringFlags >( ~(BIFF_STR_8BITLENGTH | BIFF_STR_EXTRAFONTS) ) ), "BiffObjectBase::dumpByteString - unknown flag" );
    bool b8BitLength = getFlag( nFlags, BIFF_STR_8BITLENGTH );

    OString aString = mxStrm->readByteString( !b8BitLength );
    BiffRichStringFontIdVec aFontIds;
    if( getFlag( nFlags, BIFF_STR_EXTRAFONTS ) )
        RichString::importFontIds( aFontIds, *mxStrm, false );

    // create string portions
    OUStringBuffer aBuffer;
    sal_Int32 nStrLen = aString.getLength();
    if( nStrLen > 0 )
    {
        // add leading and trailing string position to ease the following loop
        if( aFontIds.empty() || (aFontIds.front().mnPos > 0) )
            aFontIds.insert( aFontIds.begin(), BiffRichStringFontId( 0, -1 ) );
        if( aFontIds.back().mnPos < nStrLen )
            aFontIds.push_back( BiffRichStringFontId( nStrLen, -1 ) );

        // use global text encoding, if nothing special is specified
        if( eDefaultTextEnc == RTL_TEXTENCODING_DONTKNOW )
            eDefaultTextEnc = getBiffData().getTextEncoding();

        // create all string portions according to the font id vector
        for( BiffRichStringFontIdVec::const_iterator aIt = aFontIds.begin(); aIt->mnPos < nStrLen; ++aIt )
        {
            sal_Int32 nPortionLen = (aIt + 1)->mnPos - aIt->mnPos;
            if( nPortionLen > 0 )
            {
                // convert byte string to unicode string, using current font encoding
                rtl_TextEncoding eTextEnc = mxBiffData->getFontEncoding( static_cast< sal_uInt16 >( aIt->mnFontId ) );
                if( eTextEnc == RTL_TEXTENCODING_DONTKNOW )
                    eTextEnc = eDefaultTextEnc;
                aBuffer.append( OStringToOUString( aString.copy( aIt->mnPos, nPortionLen ), eTextEnc ) );
            }
        }
    }

    OUString aUniStr = aBuffer.makeStringAndClear();
    writeStringItem( pcName ? pcName : "text", aUniStr );
    return aUniStr;
}

OUString BiffObjectBase::dumpUniString( const sal_Char* pcName, BiffStringFlags nFlags )
{
    OSL_ENSURE( !getFlag( nFlags, static_cast< BiffStringFlags >( ~(BIFF_STR_8BITLENGTH | BIFF_STR_SMARTFLAGS) ) ), "BiffObjectBase::dumpUniString - unknown flag" );
    bool b8BitLength = getFlag( nFlags, BIFF_STR_8BITLENGTH );

    // --- string header ---
    sal_uInt16 nChars = b8BitLength ? mxStrm->readuInt8() : mxStrm->readuInt16();
    sal_uInt8 nFlagField = 0;
    if( (nChars > 0) || !getFlag( nFlags, BIFF_STR_SMARTFLAGS ) )
        *mxStrm >> nFlagField;

    bool b16Bit, bFonts, bPhonetic;
    sal_uInt16 nFontCount;
    sal_uInt32 nPhoneticSize;
    mxStrm->readExtendedUniStringHeader( b16Bit, bFonts, bPhonetic, nFontCount, nPhoneticSize, nFlagField );

    // --- character array ---
    OUString aString = mxStrm->readRawUniString( nChars, b16Bit );
    writeStringItem( pcName ? pcName : "text", aString );

    // --- formatting ---
    // #122185# bRich flag may be set, but format runs may be missing
    BiffRichStringFontIdVec aFontIds;
    if( nFontCount > 0 )
    {
        aFontIds.reserve( nFontCount );
        RichString::importFontIds( aFontIds, *mxStrm, nFontCount, true );
        IndentGuard aIndGuard( out() );
        writeRichStringFontIds( aFontIds );
    }

    // --- phonetic information ---
    // #122185# bPhonetic flag may be set, but phonetic data may be missing
    if( nPhoneticSize > 0 )
    {
        IndentGuard aIndGuard( out() );
        writeDecItem( "phonetic-size", nPhoneticSize );
        dumpBinary( "phonetic-data", nPhoneticSize, false );
    }

    return aString;
}

OUString BiffObjectBase::dumpString( const sal_Char* pcName, BiffStringFlags nByteFlags, BiffStringFlags nUniFlags, rtl_TextEncoding eDefaultTextEnc )
{
    return (getBiff() == BIFF8) ? dumpUniString( pcName, nUniFlags ) : dumpByteString( pcName, nByteFlags, eDefaultTextEnc );
}

OUString BiffObjectBase::dumpOleString( const sal_Char* pcName, sal_Int32 nCharCount, bool bUnicode )
{
    OUString aString;
    if( nCharCount > 0 )
    {
        sal_uInt16 nReadChars = getLimitedValue< sal_uInt16, sal_Int32 >( nCharCount, 0, SAL_MAX_UINT16 );
        aString = bUnicode ?
            mxStrm->readUnicodeArray( nReadChars ) :
            mxStrm->readCharArray( nReadChars, getBiffData().getTextEncoding() );
        // ignore remaining chars
        sal_uInt32 nIgnore = static_cast< sal_uInt32 >( nCharCount - nReadChars );
        mxStrm->ignore( bUnicode ? (nIgnore * 2) : nIgnore );
    }
    writeStringItem( pcName ? pcName : "text", aString );
    return aString;
}

OUString BiffObjectBase::dumpOleString( const sal_Char* pcName, bool bUnicode )
{
    return dumpOleString( pcName, mxStrm->readInt32(), bUnicode );
}

OUString BiffObjectBase::dumpNullString( const sal_Char* pcName, bool bUnicode )
{
    OUString aString;
    if( bUnicode )
    {
        OUStringBuffer aBuffer;
        sal_uInt16 nChar;
        for( bool bLoop = true; bLoop && mxStrm->isValid(); )
        {
            *mxStrm >> nChar;
            if( (bLoop = (nChar != 0)) == true )
                aBuffer.append( static_cast< sal_Unicode >( nChar ) );
        }
        aString = aBuffer.makeStringAndClear();
    }
    else
    {
        OStringBuffer aBuffer;
        sal_uInt8 nChar;
        for( bool bLoop = true; bLoop && mxStrm->isValid(); )
        {
            *mxStrm >> nChar;
            if( (bLoop = (nChar != 0)) == true )
                aBuffer.append( static_cast< sal_Char >( nChar ) );
        }
        aString = OStringToOUString( aBuffer.makeStringAndClear(), getBiffData().getTextEncoding() );
    }
    writeStringItem( pcName ? pcName : "text", aString );
    return aString;
}

double BiffObjectBase::dumpRk( const sal_Char* pcName )
{
    sal_Int32 nRk;
    *mxStrm >> nRk;
    return writeRkItem( pcName ? pcName : "rk-value", nRk );
}

sal_uInt8 BiffObjectBase::dumpBoolean( const sal_Char* pcName )
{
    sal_uInt8 nBool;
    *mxStrm >> nBool;
    writeBooleanItem( pcName ? pcName : "boolean", nBool );
    return nBool;
}

sal_uInt8 BiffObjectBase::dumpErrorCode( const sal_Char* pcName )
{
    sal_uInt8 nErrCode;
    *mxStrm >> nErrCode;
    writeErrorCodeItem( pcName ? pcName : "errorcode", nErrCode );
    return nErrCode;
}

sal_Int32 BiffObjectBase::dumpRgbColor( const sal_Char* pcName )
{
    OoxColor aColor;
    aColor.importColorRgb( *mxStrm );
    writeColorItem( pcName ? pcName : "color-rgb", aColor.mnValue );
    return aColor.mnValue;
}

rtl_TextEncoding BiffObjectBase::dumpCodePage( const sal_Char* pcName )
{
    sal_uInt16 nCodePage = dumpDec< sal_uInt16 >( pcName ? pcName : "codepage", "CODEPAGES" );
    return BiffHelper::calcTextEncodingFromCodePage( nCodePage );
}

void BiffObjectBase::dumpFormulaResult( const sal_Char* pcName )
{
    MultiItemsGuard aMultiGuard( out() );
    sal_uInt8 pnResult[ 8 ];
    mxStrm->read( pnResult, 8 );
    writeArrayItem( pcName ? pcName : "result", pnResult, 8 );
    if( (pnResult[ 6 ] == 0xFF) && (pnResult[ 7 ] == 0xFF) )
    {
        sal_uInt8 nType = pnResult[ 0 ];
        sal_uInt8 nData = pnResult[ 2 ];
        writeHexItem( "type", nType, mxResultType );
        switch( nType )
        {
            case BIFF_FORMULA_RES_BOOL:     writeBooleanItem( "value", nData );     break;
            case BIFF_FORMULA_RES_ERROR:    writeErrorCodeItem( "value", nData );   break;
        }
    }
    else
    {
        double* pfValue = reinterpret_cast< double* >( pnResult );
        ::oox::core::ByteOrderConverter::convertLittleEndian( *pfValue );
        writeDecItem( "value", *pfValue );
    }
}

sal_Int32 BiffObjectBase::dumpColIndex( const sal_Char* pcName, bool bCol16Bit )
{
    sal_Int32 nCol = lclReadCol( *mxStrm, bCol16Bit );
    writeColIndexItem( pcName ? pcName : "col-idx", nCol );
    return nCol;
}

sal_Int32 BiffObjectBase::dumpColRange( const sal_Char* pcName, bool bCol16Bit )
{
    sal_Int32 nCol1 = lclReadCol( *mxStrm, bCol16Bit );
    sal_Int32 nCol2 = lclReadCol( *mxStrm, bCol16Bit );
    writeColRangeItem( pcName ? pcName : "col-range", nCol1, nCol2 );
    return nCol2 - nCol1 + 1;
}

sal_Int32 BiffObjectBase::dumpRowIndex( const sal_Char* pcName, bool bRow32Bit )
{
    sal_Int32 nRow = lclReadRow( *mxStrm, bRow32Bit );
    writeRowIndexItem( pcName ? pcName : "row-idx", nRow );
    return nRow;
}

sal_Int32 BiffObjectBase::dumpRowRange( const sal_Char* pcName, bool bRow32Bit )
{
    sal_Int32 nRow1 = lclReadRow( *mxStrm, bRow32Bit );
    sal_Int32 nRow2 = lclReadRow( *mxStrm, bRow32Bit );
    writeRowRangeItem( pcName ? pcName : "row-range", nRow1, nRow2 );
    return nRow2 - nRow1 + 1;
}

Address BiffObjectBase::dumpAddress( const sal_Char* pcName, bool bCol16Bit, bool bRow32Bit )
{
    Address aPos;
    aPos.read( *mxStrm, bCol16Bit, bRow32Bit );
    writeAddressItem( pcName ? pcName : "addr", aPos );
    return aPos;
}

Range BiffObjectBase::dumpRange( const sal_Char* pcName, bool bCol16Bit, bool bRow32Bit )
{
    Range aRange;
    aRange.read( *mxStrm, bCol16Bit, bRow32Bit );
    writeRangeItem( pcName ? pcName : "range", aRange );
    return aRange;
}

void BiffObjectBase::dumpRangeList( const sal_Char* pcName, bool bCol16Bit, bool bRow32Bit )
{
    RangeList aRanges;
    aRanges.read( *mxStrm, bCol16Bit, bRow32Bit );
    writeRangeListItem( pcName ? pcName : "range-list", aRanges );
}

void BiffObjectBase::dumpConstArrayHeader( sal_uInt32& rnCols, sal_uInt32& rnRows )
{
    Output& rOut = out();
    MultiItemsGuard aMultiGuard( rOut );
    rnCols = dumpDec< sal_uInt8 >( "width" );
    rnRows = dumpDec< sal_uInt16 >( "height" );
    switch( getBiff() )
    {
        case BIFF2:
        case BIFF3:
        case BIFF4:
        case BIFF5: if( rnCols == 0 ) rnCols = 256; break;
        case BIFF8: ++rnCols; ++rnRows;             break;
        case BIFF_UNKNOWN:                          break;
    }
    ItemGuard aItem( rOut, "size" );
    rOut.writeDec( rnCols );
    rOut.writeChar( 'x' );
    rOut.writeDec( rnRows );
    aItem.cont();
    rOut.writeDec( rnCols * rnRows );
}

OUString BiffObjectBase::dumpConstValue()
{
    Output& rOut = out();
    MultiItemsGuard aMultiGuard( rOut );
    OUStringBuffer aValue;
    switch( dumpDec< sal_uInt8 >( "type", mxConstType ) )
    {
        case BIFF_DATATYPE_EMPTY:
            dumpUnused( 8 );
            aValue.append( OOX_DUMP_EMPTYVALUE );
        break;
        case BIFF_DATATYPE_DOUBLE:
            dumpDec< double >( "value" );
            aValue.append( rOut.getLastItemValue() );
        break;
        case BIFF_DATATYPE_STRING:
            aValue.append( dumpString( "value", BIFF_STR_8BITLENGTH ) );
            StringHelper::enclose( aValue, '"' );
        break;
        case BIFF_DATATYPE_BOOL:
            dumpBoolean( "value" );
            aValue.append( rOut.getLastItemValue() );
            dumpUnused( 7 );
        break;
        case BIFF_DATATYPE_ERROR:
            dumpErrorCode( "value" );
            aValue.append( rOut.getLastItemValue() );
            dumpUnused( 7 );
        break;
    }
    return aValue.makeStringAndClear();
}

// ----------------------------------------------------------------------------

void BiffObjectBase::constructBiffObjBase( BiffType eBiff )
{
    if( WrappedStreamObject::implIsValid() )
    {
        BinaryInputStream& rStrm = getStream();
        mxBiffData.reset( new BiffSharedData(
            (eBiff == BIFF_UNKNOWN) ? BiffFormatDetector::detectStreamBiffVersion( rStrm ) : eBiff ) );
        mxStrm.reset( new BiffInputStream( rStrm ) );
        reconstructConfig();
        reconstructInput();
    }
    if( BiffObjectBase::implIsValid() )
    {
        const Config& rCfg = cfg();
        mxBoolean = rCfg.getNameList( "BOOLEAN" );
        mxErrCodes = rCfg.getNameList( "ERRORCODES" );
        mxConstType = rCfg.getNameList( "CONSTVALUE-TYPE" );
        mxResultType = rCfg.getNameList( "FORMULA-RESULTTYPE" );
    }
}

// ============================================================================

SimpleObjectBase::~SimpleObjectBase()
{
}

void SimpleObjectBase::construct( const BiffObjectBase& rParent )
{
    BiffObjectBase::construct( rParent );
}

void SimpleObjectBase::implDumpHeader()
{
}

void SimpleObjectBase::implDumpFooter()
{
}

// ============================================================================
// ============================================================================

FormulaStack::FormulaStack() :
    mbError( false )
{
}

void FormulaStack::pushOperand( const StringWrapper& rOp, const OUString& rTokClass )
{
    maFmlaStack.push( rOp.getString() );
    maClassStack.push( rTokClass );
}

void FormulaStack::pushOperand( const StringWrapper& rOp )
{
    pushOperand( rOp, OUString( OOX_DUMP_BASECLASS ) );
}

void FormulaStack::pushUnaryOp( const StringWrapper& rLOp, const StringWrapper& rROp )
{
    pushUnaryOp( maFmlaStack, rLOp.getString(), rROp.getString() );
    pushUnaryOp( maClassStack, rLOp.getString(), rROp.getString() );
}

void FormulaStack::pushBinaryOp( const StringWrapper& rOp )
{
    pushBinaryOp( maFmlaStack, rOp.getString() );
    pushBinaryOp( maClassStack, rOp.getString() );
}

void FormulaStack::pushFuncOp( const StringWrapper& rFunc, const OUString& rTokClass, sal_uInt8 nParamCount )
{
    pushFuncOp( maFmlaStack, rFunc.getString(), nParamCount );
    pushFuncOp( maClassStack, rTokClass, nParamCount );
}

void FormulaStack::replaceOnTop( const OUString& rOld, const OUString& rNew )
{
    if( !maFmlaStack.empty() )
    {
        sal_Int32 nPos = maFmlaStack.top().indexOf( rOld );
        if( nPos >= 0 )
            maFmlaStack.top() = maFmlaStack.top().copy( 0, nPos ) + rNew + maFmlaStack.top().copy( nPos + rOld.getLength() );
    }
}

const OUString& FormulaStack::getString( const StringStack& rStack ) const
{
    static const OUString saStackError = OOX_DUMP_ERRSTRING( "stack" );
    return (mbError || rStack.empty()) ? saStackError : rStack.top();
}

void FormulaStack::pushUnaryOp( StringStack& rStack, const OUString& rLOp, const OUString& rROp )
{
    if( check( !rStack.empty() ) )
        rStack.top() = rLOp + rStack.top() + rROp;
}

void FormulaStack::pushBinaryOp( StringStack& rStack, const OUString& rOp )
{
    OUString aSecond;
    if( check( !rStack.empty() ) )
    {
        aSecond = rStack.top();
        rStack.pop();
    }
    if( check( !rStack.empty() ) )
        rStack.top() = rStack.top() + rOp + aSecond;
}

void FormulaStack::pushFuncOp( StringStack& rStack, const OUString& rOp, sal_uInt8 nParamCount )
{
    OUStringBuffer aFunc;
    for( sal_uInt8 nParam = 0; (nParam < nParamCount) && check( !rStack.empty() ); ++nParam )
    {
        StringHelper::prependToken( aFunc, rStack.top(), OOX_DUMP_FUNCSEP );
        rStack.pop();
    }
    StringHelper::enclose( aFunc, '(', ')' );
    aFunc.insert( 0, rOp );
    rStack.push( aFunc.makeStringAndClear() );
}

// ============================================================================

FormulaObject::FormulaObject( const BiffObjectBase& rParent ) :
    mpcName( 0 ),
    mnSize( 0 )
{
    BiffObjectBase::construct( rParent );
    constructFmlaObj();
}

FormulaObject::~FormulaObject()
{
}

sal_uInt16 FormulaObject::readFormulaSize()
{
    return (getBiff() == BIFF2) ? getBiffStream().readuInt8() : getBiffStream().readuInt16();
}

sal_uInt16 FormulaObject::dumpFormulaSize( const sal_Char* pcName )
{
    if( !pcName ) pcName = "formula-size";
    sal_uInt16 nSize = readFormulaSize();
    writeDecItem( pcName, nSize );
    return nSize;
}

void FormulaObject::dumpCellFormula( const sal_Char* pcName, sal_uInt16 nSize )
{
    dumpFormula( pcName, nSize, false );
}

void FormulaObject::dumpCellFormula( const sal_Char* pcName )
{
    dumpFormula( pcName, false );
}

void FormulaObject::dumpNameFormula( const sal_Char* pcName, sal_uInt16 nSize )
{
    dumpFormula( pcName, nSize, true );
}

void FormulaObject::dumpNameFormula( const sal_Char* pcName )
{
    dumpFormula( pcName, true );
}

void FormulaObject::implDumpHeader()
{
    MultiItemsGuard aMultiGuard( out() );
    writeEmptyItem( mpcName );
    writeDecItem( "formula-size", mnSize );
}

void FormulaObject::implDumpBody()
{
    if( mnSize == 0 ) return;

    Input& rIn = in();
    sal_Int64 nStartPos = rIn.tell();
    sal_Int64 nEndPos = ::std::min< sal_Int64 >( nStartPos + mnSize, rIn.getSize() );

    bool bValid = mxTokens.get();
    mxStack.reset( new FormulaStack );
    maAddData.clear();
    IndentGuard aIndGuard( out() );
    {
        TableGuard aTabGuard( out(), 8, 18 );
        while( bValid && (rIn.tell() < nEndPos) )
        {
            MultiItemsGuard aMultiGuard( out() );
            writeHexItem( 0, static_cast< sal_uInt16 >( rIn.tell() - nStartPos ) );
            sal_uInt8 nTokenId = dumpHex< sal_uInt8 >( 0, mxTokens );
            bValid = mxTokens->hasName( nTokenId );
            if( bValid )
            {
                sal_uInt8 nTokClass = nTokenId & BIFF_TOKCLASS_MASK;
                sal_uInt8 nBaseId = nTokenId & BIFF_TOKID_MASK;
                if( nTokClass == BIFF_TOKCLASS_NONE )
                {
                    switch( nBaseId )
                    {
                        case BIFF_TOKID_EXP:        dumpExpToken( "EXP" );          break;
                        case BIFF_TOKID_TBL:        dumpExpToken( "TBL" );          break;
                        case BIFF_TOKID_ADD:        dumpBinaryOpToken( "+" );       break;
                        case BIFF_TOKID_SUB:        dumpBinaryOpToken( "-" );       break;
                        case BIFF_TOKID_MUL:        dumpBinaryOpToken( "*" );       break;
                        case BIFF_TOKID_DIV:        dumpBinaryOpToken( "/" );       break;
                        case BIFF_TOKID_POWER:      dumpBinaryOpToken( "^" );       break;
                        case BIFF_TOKID_CONCAT:     dumpBinaryOpToken( "&" );       break;
                        case BIFF_TOKID_LT:         dumpBinaryOpToken( "<" );       break;
                        case BIFF_TOKID_LE:         dumpBinaryOpToken( "<=" );      break;
                        case BIFF_TOKID_EQ:         dumpBinaryOpToken( "=" );       break;
                        case BIFF_TOKID_GE:         dumpBinaryOpToken( ">=" );      break;
                        case BIFF_TOKID_GT:         dumpBinaryOpToken( "<" );       break;
                        case BIFF_TOKID_NE:         dumpBinaryOpToken( "<>" );      break;
                        case BIFF_TOKID_ISECT:      dumpBinaryOpToken( " " );       break;
                        case BIFF_TOKID_LIST:       dumpBinaryOpToken( "," );       break;
                        case BIFF_TOKID_RANGE:      dumpBinaryOpToken( ":" );       break;
                        case BIFF_TOKID_UPLUS:      dumpUnaryOpToken( "+", "" );    break;
                        case BIFF_TOKID_UMINUS:     dumpUnaryOpToken( "-", "" );    break;
                        case BIFF_TOKID_PERCENT:    dumpUnaryOpToken( "", "%" );    break;
                        case BIFF_TOKID_PAREN:      dumpUnaryOpToken( "(", ")" );   break;
                        case BIFF_TOKID_MISSARG:    dumpMissArgToken();             break;
                        case BIFF_TOKID_STR:        dumpStringToken();              break;
                        case BIFF_TOKID_NLR:        bValid = dumpNlrToken();        break;
                        case BIFF_TOKID_ATTR:       bValid = dumpAttrToken();       break;
                        case BIFF_TOKID_SHEET:      dumpSheetToken();               break;
                        case BIFF_TOKID_ENDSHEET:   dumpEndSheetToken();            break;
                        case BIFF_TOKID_ERR:        dumpErrorToken();               break;
                        case BIFF_TOKID_BOOL:       dumpBoolToken();                break;
                        case BIFF_TOKID_INT:        dumpIntToken();                 break;
                        case BIFF_TOKID_NUM:        dumpDoubleToken();              break;
                    }
                }
                else
                {
                    OUString aTokClass = cfg().getName( mxClasses, nTokClass );
                    switch( nBaseId )
                    {
                        case BIFF_TOKID_ARRAY:      dumpArrayToken( aTokClass );                break;
                        case BIFF_TOKID_FUNC:       dumpFuncToken( aTokClass );                 break;
                        case BIFF_TOKID_FUNCVAR:    dumpFuncVarToken( aTokClass );              break;
                        case BIFF_TOKID_NAME:       dumpNameToken( aTokClass );                 break;
                        case BIFF_TOKID_REF:        dumpRefToken( aTokClass, false );           break;
                        case BIFF_TOKID_AREA:       dumpAreaToken( aTokClass, false );          break;
                        case BIFF_TOKID_MEMAREA:    dumpMemAreaToken( aTokClass, true );        break;
                        case BIFF_TOKID_MEMERR:     dumpMemAreaToken( aTokClass, false );       break;
                        case BIFF_TOKID_MEMNOMEM:   dumpMemAreaToken( aTokClass, false );       break;
                        case BIFF_TOKID_MEMFUNC:    dumpMemFuncToken( aTokClass );              break;
                        case BIFF_TOKID_REFERR:     dumpRefErrToken( aTokClass, false );        break;
                        case BIFF_TOKID_AREAERR:    dumpRefErrToken( aTokClass, true );         break;
                        case BIFF_TOKID_REFN:       dumpRefToken( aTokClass, true );            break;
                        case BIFF_TOKID_AREAN:      dumpAreaToken( aTokClass, true );           break;
                        case BIFF_TOKID_MEMAREAN:   dumpMemFuncToken( aTokClass );              break;
                        case BIFF_TOKID_MEMNOMEMN:  dumpMemFuncToken( aTokClass );              break;
                        case BIFF_TOKID_FUNCCE:     dumpCmdToken( aTokClass );                  break;
                        case BIFF_TOKID_NAMEX:      dumpNameXToken( aTokClass );                break;
                        case BIFF_TOKID_REF3D:      dumpRef3dToken( aTokClass, mbNameMode );    break;
                        case BIFF_TOKID_AREA3D:     dumpArea3dToken( aTokClass, mbNameMode );   break;
                        case BIFF_TOKID_REFERR3D:   dumpRefErr3dToken( aTokClass, false );      break;
                        case BIFF_TOKID_AREAERR3D:  dumpRefErr3dToken( aTokClass, true );       break;
                    }
                }
            }
        }
    }
    bValid = nEndPos == rIn.tell();
    if( bValid )
    {
        dumpAddTokenData();
        writeInfoItem( "formula", mxStack->getFormulaString() );
        writeInfoItem( "classes", mxStack->getClassesString() );
    }
    else
        dumpBinary( OOX_DUMP_ERRASCII( "formula-error" ), static_cast< sal_Int32 >( nEndPos - rIn.tell() ), false );

    mpcName = 0;
    mnSize = 0;
}

void FormulaObject::dumpFormula( const sal_Char* pcName, sal_uInt16 nSize, bool bNameMode )
{
    mpcName = pcName ? pcName : "formula";
    mnSize = nSize;
    mbNameMode = bNameMode;
    dump();
}

void FormulaObject::dumpFormula( const sal_Char* pcName, bool bNameMode )
{
    dumpFormula( pcName, readFormulaSize(), bNameMode );
}

// private --------------------------------------------------------------------

void FormulaObject::constructFmlaObj()
{
    if( BiffObjectBase::implIsValid() )
    {
        Reference< XSpreadsheetDocument > xDocument( getFilter().getModel(), UNO_QUERY );
        mxFuncProv.reset( new FunctionProvider( xDocument, getBiff(), true ) );

        Config& rCfg = cfg();
        mxClasses   = rCfg.getNameList( "TOKENCLASSES" );
        mxFuncNames = rCfg.getNameList( "FUNCTIONNAMES" );
        mxParamCnt  = rCfg.getNameList( "PARAMCOUNT" );
        mxRelFlags  = rCfg.getNameList( "REFRELFLAGS" );
        mxNlrTypes  = rCfg.getNameList( "NLRTYPES" );
        mxAttrTypes = rCfg.getNameList( "ATTRTYPES" );
        mxSpTypes   = rCfg.getNameList( "ATTRSPACETYPES" );

        // create classified token names
        mxTokens = rCfg.createNameList< ConstList >( "TOKENS" );
        mxTokens->includeList( rCfg.getNameList( "BASETOKENS" ) );

        NameListRef xClassTokens = rCfg.getNameList( "CLASSTOKENS" );
        if( mxClasses.get() && xClassTokens.get() )
            for( NameListBase::const_iterator aCIt = mxClasses->begin(), aCEnd = mxClasses->end(); aCIt != aCEnd; ++aCIt )
                for( NameListBase::const_iterator aTIt = xClassTokens->begin(), aTEnd = xClassTokens->end(); aTIt != aTEnd; ++aTIt )
                    mxTokens->setName( aCIt->first | aTIt->first, aTIt->second + aCIt->second );

        mnColCount = 256;
        mnRowCount = (getBiff() == BIFF8) ? 65536 : 16384;
    }
}

// ----------------------------------------------------------------------------

namespace {

OUString lclCreateName( const OUString& rRef, sal_uInt16 nNameIdx )
{
    OUStringBuffer aName( rRef );
    StringHelper::appendIndexedText( aName, CREATE_OUSTRING( "NAME" ), nNameIdx );
    return aName.makeStringAndClear();
}

OUString lclCreateNlr( const OUString& rData, bool bRel = true )
{
    OUStringBuffer aNlr;
    if( !bRel ) aNlr.append( OOX_DUMP_ADDRABS );
    StringHelper::appendIndexedText( aNlr, CREATE_OUSTRING( "NLR" ), rData );
    return aNlr.makeStringAndClear();
}

OUString lclCreateNlr( const TokenAddress& rPos )
{
    OUStringBuffer aAddr;
    StringHelper::appendAddrCol( aAddr, rPos.mnCol, true );
    StringHelper::appendAddrRow( aAddr, rPos.mnRow, true );
    return lclCreateNlr( aAddr.makeStringAndClear(), rPos.mbRelRow );
}

} // namespace

// ----------------------------------------------------------------------------

TokenAddress FormulaObject::createTokenAddress( sal_uInt16 nCol, sal_uInt16 nRow, bool bRelC, bool bRelR, bool bNameMode ) const
{
    TokenAddress aPos;
    aPos.mnCol = nCol;
    if( bRelC && bNameMode && (nCol >= mnColCount / 2) ) aPos.mnCol -= mnColCount;
    aPos.mbRelCol = bRelC;
    aPos.mnRow = nRow;
    if( bRelR && bNameMode && (nRow >= mnRowCount / 2) ) aPos.mnRow -= mnRowCount;
    aPos.mbRelRow = bRelR;
    return aPos;
}

OUString FormulaObject::createFunc( sal_uInt16 nFuncIdx ) const
{
    OUStringBuffer aFunc;
    if( cfg().hasName( mxFuncNames, nFuncIdx ) )
        aFunc.append( cfg().getName( mxFuncNames, nFuncIdx ) );
    else
    {
        bool bCmd = getFlag( nFuncIdx, BIFF_TOK_FUNCVAR_CMD );
        aFunc.appendAscii( bCmd ? "CMD" : "FUNC" );
        StringHelper::appendIndex( aFunc, nFuncIdx & BIFF_TOK_FUNCVAR_INDEXMASK );
    }
    return aFunc.makeStringAndClear();
}

OUString FormulaObject::createRef( const OUString& rData ) const
{
    return maRefPrefix + rData;
}

OUString FormulaObject::createName( sal_uInt16 nNameIdx ) const
{
    return lclCreateName( maRefPrefix, nNameIdx );
}

OUString FormulaObject::createPlaceHolder( size_t nIdx ) const
{
    OUStringBuffer aStr;
    StringHelper::appendDec( aStr, static_cast< sal_uInt32 >( nIdx ) );
    StringHelper::enclose( aStr, OOX_DUMP_PLACEHOLDER );
    return aStr.makeStringAndClear();
}

OUString FormulaObject::createPlaceHolder() const
{
    return createPlaceHolder( maAddData.size() );
}

void FormulaObject::writeTokenAddressItem( const sal_Char* pcName, const TokenAddress& rPos, bool bNameMode )
{
    ItemGuard aItem( out(), pcName );
    StringHelper::appendAddress( out().getLine(), rPos, bNameMode );
}

void FormulaObject::writeTokenAddress3dItem( const sal_Char* pcName, const OUString& rRef, const TokenAddress& rPos, bool bNameMode )
{
    ItemGuard aItem( out(), pcName );
    out().writeString( rRef );
    StringHelper::appendAddress( out().getLine(), rPos, bNameMode );
}

void FormulaObject::writeTokenRangeItem( const sal_Char* pcName, const TokenRange& rRange, bool bNameMode )
{
    ItemGuard aItem( out(), pcName );
    StringHelper::appendRange( out().getLine(), rRange, bNameMode );
}

void FormulaObject::writeTokenRange3dItem( const sal_Char* pcName, const OUString& rRef, const TokenRange& rRange, bool bNameMode )
{
    ItemGuard aItem( out(), pcName );
    out().writeString( rRef );
    StringHelper::appendRange( out().getLine(), rRange, bNameMode );
}

sal_uInt16 FormulaObject::dumpTokenFuncIdx()
{
    return dumpDec< sal_uInt16, sal_uInt8 >( getBiff() >= BIFF4, "func-idx", mxFuncNames );
}

sal_uInt16 FormulaObject::dumpTokenCol( const sal_Char* pcName, bool& rbRelC, bool& rbRelR )
{
    sal_uInt16 nCol = 0;
    if( getBiff() == BIFF8 )
    {
        nCol = dumpHex< sal_uInt16 >( pcName, mxRelFlags );
        rbRelC = getFlag( nCol, BIFF_TOK_REF_COLREL );
        rbRelR = getFlag( nCol, BIFF_TOK_REF_ROWREL );
        nCol &= BIFF_TOK_REF_COLMASK;
    }
    else
        nCol = dumpDec< sal_uInt8 >( pcName );
    return nCol;
}

sal_uInt16 FormulaObject::dumpTokenRow( const sal_Char* pcName, bool& rbRelC, bool& rbRelR )
{
    sal_uInt16 nRow = 0;
    if( getBiff() == BIFF8 )
        nRow = dumpDec< sal_uInt16 >( pcName );
    else
    {
        nRow = dumpHex< sal_uInt16 >( pcName, mxRelFlags );
        rbRelC = getFlag( nRow, BIFF_TOK_REF_COLREL );
        rbRelR = getFlag( nRow, BIFF_TOK_REF_ROWREL );
        nRow &= BIFF_TOK_REF_ROWMASK;
    }
    return nRow;
}

TokenAddress FormulaObject::dumpTokenAddress( bool bNameMode )
{
    bool bRelC = false;
    bool bRelR = false;
    sal_uInt16 nRow = dumpTokenRow( "row", bRelC, bRelR );
    sal_uInt16 nCol = dumpTokenCol( "col", bRelC, bRelR );
    return createTokenAddress( nCol, nRow, bRelC, bRelR, bNameMode );
}

TokenRange FormulaObject::dumpTokenRange( bool bNameMode )
{
    bool bRelC1 = false;
    bool bRelR1 = false;
    bool bRelC2 = false;
    bool bRelR2 = false;
    sal_uInt16 nRow1 = dumpTokenRow( "row1", bRelC1, bRelR1 );
    sal_uInt16 nRow2 = dumpTokenRow( "row2", bRelC2, bRelR2 );
    sal_uInt16 nCol1 = dumpTokenCol( "col1", bRelC1, bRelR1 );
    sal_uInt16 nCol2 = dumpTokenCol( "col2", bRelC2, bRelR2 );
    TokenRange aRange;
    aRange.maFirst = createTokenAddress( nCol1, nRow1, bRelC1, bRelR1, bNameMode );
    aRange.maLast  = createTokenAddress( nCol2, nRow2, bRelC2, bRelR2, bNameMode );
    return aRange;
}

OUString FormulaObject::dumpTokenRefIdx()
{
    sal_Int16 nRefIdx = dumpDec< sal_Int16 >( "ref-idx" );
    switch( getBiff() )
    {
        case BIFF2: dumpUnused( 1 );    break;
        case BIFF3: dumpUnused( 2 );    break;
        case BIFF4: dumpUnused( 2 );    break;
        case BIFF5: dumpUnused( 8 );    break;
        case BIFF8:                     break;
        case BIFF_UNKNOWN:              break;
    }
    OUStringBuffer aRef( CREATE_OUSTRING( "REF" ) );
    StringHelper::appendIndex( aRef, nRefIdx );
    aRef.append( OOX_DUMP_TABSEP );
    return aRef.makeStringAndClear();
}

void FormulaObject::dumpTokenRefTabIdxs()
{
    if( getBiff() <= BIFF5 )
    {
        dumpDec< sal_Int16 >( "tab1" );
        dumpDec< sal_Int16 >( "tab2" );
    }
}

void FormulaObject::dumpIntToken()
{
    dumpDec< sal_uInt16 >( "value" );
    mxStack->pushOperand( out().getLastItemValue() );
}

void FormulaObject::dumpDoubleToken()
{
    dumpDec< double >( "value" );
    mxStack->pushOperand( out().getLastItemValue() );
}

void FormulaObject::dumpStringToken()
{
    OUStringBuffer aValue;
    aValue.append( dumpString( "value", BIFF_STR_8BITLENGTH, BIFF_STR_8BITLENGTH ) );
    StringHelper::enclose( aValue, '"' );
    mxStack->pushOperand( aValue.makeStringAndClear() );
}

void FormulaObject::dumpBoolToken()
{
    dumpBoolean( "value" );
    mxStack->pushOperand( out().getLastItemValue() );
}

void FormulaObject::dumpErrorToken()
{
    dumpErrorCode( "value" );
    mxStack->pushOperand( out().getLastItemValue() );
}

void FormulaObject::dumpMissArgToken()
{
    mxStack->pushOperand( OUString( OOX_DUMP_EMPTYVALUE ) );
}

void FormulaObject::dumpArrayToken( const OUString& rTokClass )
{
    dumpUnused( (getBiff() == BIFF2) ? 6 : 7 );
    mxStack->pushOperand( createPlaceHolder(), rTokClass );
    maAddData.push_back( ADDDATA_ARRAY );
}

void FormulaObject::dumpNameToken( const OUString& rTokClass )
{
    sal_uInt16 nNameIdx = dumpDec< sal_uInt16 >( "name-idx" );
    switch( getBiff() )
    {
        case BIFF2: dumpUnused( 5 );    break;
        case BIFF3:
        case BIFF4: dumpUnused( 8 );    break;
        case BIFF5: dumpUnused( 12 );   break;
        case BIFF8: dumpUnused( 2 );    break;
        case BIFF_UNKNOWN:              break;
    }
    mxStack->pushOperand( createName( nNameIdx ), rTokClass );
}

void FormulaObject::dumpNameXToken( const OUString& rTokClass )
{
    OUString aRef = dumpTokenRefIdx();
    sal_uInt16 nNameIdx = dumpDec< sal_uInt16 >( "name-idx" );
    dumpUnused( (getBiff() == BIFF8) ? 2 : 12 );
    mxStack->pushOperand( lclCreateName( aRef, nNameIdx ), rTokClass );
}

void FormulaObject::dumpRefToken( const OUString& rTokClass, bool bNameMode )
{
    TokenAddress aPos = dumpTokenAddress( bNameMode );
    writeTokenAddressItem( "addr", aPos, bNameMode );
    mxStack->pushOperand( createRef( out().getLastItemValue() ), rTokClass );
}

void FormulaObject::dumpAreaToken( const OUString& rTokClass, bool bNameMode )
{
    TokenRange aRange = dumpTokenRange( bNameMode );
    writeTokenRangeItem( "range", aRange, bNameMode );
    mxStack->pushOperand( createRef( out().getLastItemValue() ), rTokClass );
}

void FormulaObject::dumpRefErrToken( const OUString& rTokClass, bool bArea )
{
    dumpUnused( ((getBiff() == BIFF8) ? 4 : 3) * (bArea ? 2 : 1) );
    mxStack->pushOperand( createRef( getErrorName( BIFF_ERR_REF ) ), rTokClass );
}

void FormulaObject::dumpRef3dToken( const OUString& rTokClass, bool bNameMode )
{
    OUString aRef = dumpTokenRefIdx();
    dumpTokenRefTabIdxs();
    TokenAddress aPos = dumpTokenAddress( bNameMode );
    writeTokenAddress3dItem( "addr", aRef, aPos, bNameMode );
    mxStack->pushOperand( out().getLastItemValue(), rTokClass );
}

void FormulaObject::dumpArea3dToken( const OUString& rTokClass, bool bNameMode )
{
    OUString aRef = dumpTokenRefIdx();
    dumpTokenRefTabIdxs();
    TokenRange aRange = dumpTokenRange( bNameMode );
    writeTokenRange3dItem( "range", aRef, aRange, bNameMode );
    mxStack->pushOperand( out().getLastItemValue(), rTokClass );
}

void FormulaObject::dumpRefErr3dToken( const OUString& rTokClass, bool bArea )
{
    OUString aRef = dumpTokenRefIdx();
    dumpTokenRefTabIdxs();
    dumpUnused( ((getBiff() == BIFF8) ? 4 : 3) * (bArea ? 2 : 1) );
    mxStack->pushOperand( aRef + getErrorName( BIFF_ERR_REF ), rTokClass );
}

void FormulaObject::dumpMemFuncToken( const OUString& /*rTokClass*/ )
{
    dumpDec< sal_uInt16, sal_uInt8 >( getBiff() != BIFF2, "size" );
}

void FormulaObject::dumpMemAreaToken( const OUString& rTokClass, bool bAddData )
{
    dumpUnused( (getBiff() == BIFF2) ? 3 : 4 );
    dumpMemFuncToken( rTokClass );
    if( bAddData )
        maAddData.push_back( ADDDATA_MEMAREA );
}

void FormulaObject::dumpExpToken( const StringWrapper& rName )
{
    Address aPos;
    aPos.mnRow = dumpDec< sal_uInt16 >( "row" );
    aPos.mnCol = dumpDec< sal_uInt16, sal_uInt8 >( getBiff() != BIFF2, "col" );
    writeAddressItem( "base-addr", aPos );
    OUStringBuffer aOp( rName.getString() );
    StringHelper::appendIndex( aOp, out().getLastItemValue() );
    mxStack->pushOperand( aOp.makeStringAndClear() );
}

void FormulaObject::dumpUnaryOpToken( const StringWrapper& rLOp, const StringWrapper& rROp )
{
    mxStack->pushUnaryOp( rLOp, rROp );
}

void FormulaObject::dumpBinaryOpToken( const StringWrapper& rOp )
{
    mxStack->pushBinaryOp( rOp );
}

void FormulaObject::dumpFuncToken( const OUString& rTokClass )
{
    sal_uInt16 nFuncIdx = dumpTokenFuncIdx();
    OUString aFuncName = createFunc( nFuncIdx );
    if( const FunctionInfo* pFuncInfo = mxFuncProv->getFuncInfoFromBiffFuncId( nFuncIdx ) )
        mxStack->pushFuncOp( aFuncName, rTokClass, pFuncInfo->mnMinParamCount );
    else
        mxStack->setError();
}

void FormulaObject::dumpFuncVarToken( const OUString& rTokClass )
{
    sal_uInt8 nParamCount = dumpDec< sal_uInt8 >( "param-count", mxParamCnt );
    bool bPrompt = getFlag( nParamCount, BIFF_TOK_FUNCVAR_PROMPT );
    nParamCount &= BIFF_TOK_FUNCVAR_COUNTMASK;
    sal_uInt16 nFuncIdx = dumpTokenFuncIdx();
    OUString aFuncName = createFunc( nFuncIdx );
    if( bPrompt ) aFuncName = aFuncName + OOX_DUMP_CMDPROMPT;
    mxStack->pushFuncOp( aFuncName, rTokClass, nParamCount );
}

void FormulaObject::dumpCmdToken( const OUString& rTokClass )
{
    sal_uInt8 nParamCount = dumpDec< sal_uInt8 >( "param-count" );
    sal_uInt16 nCmdIdx = dumpTokenFuncIdx() | BIFF_TOK_FUNCVAR_CMD;
    mxStack->pushFuncOp( createFunc( nCmdIdx ), rTokClass, nParamCount );
}

void FormulaObject::dumpSheetToken()
{
    dumpUnused( (getBiff() == BIFF2) ? 4 : 6 );
    maRefPrefix = dumpTokenRefIdx();
}

void FormulaObject::dumpEndSheetToken()
{
    dumpUnused( (getBiff() == BIFF2) ? 3 : 4 );
    maRefPrefix = OUString();
}

bool FormulaObject::dumpAttrToken()
{
    bool bValid = true;
    bool bBiff2 = getBiff() == BIFF2;
    sal_uInt8 nType = dumpHex< sal_uInt8 >( "type", mxAttrTypes );
    switch( nType )
    {
        case BIFF_TOK_ATTR_VOLATILE:
            dumpUnused( bBiff2 ? 1 : 2 );
        break;
        case BIFF_TOK_ATTR_IF:
            dumpDec< sal_uInt16, sal_uInt8 >( !bBiff2, "skip" );
        break;
        case BIFF_TOK_ATTR_CHOOSE:
        {
            sal_uInt16 nCount = dumpDec< sal_uInt16, sal_uInt8 >( !bBiff2, "choices" );
            out().resetItemIndex();
            for( sal_uInt16 nIdx = 0; nIdx < nCount; ++nIdx )
                dumpDec< sal_uInt16, sal_uInt8 >( !bBiff2, "#skip" );
            dumpDec< sal_uInt16, sal_uInt8 >( !bBiff2, "skip-err" );
        }
        break;
        case BIFF_TOK_ATTR_SKIP:
            dumpDec< sal_uInt16, sal_uInt8 >( !bBiff2, "skip" );
        break;
        case BIFF_TOK_ATTR_SUM:
            dumpUnused( bBiff2 ? 1 : 2 );
            mxStack->pushFuncOp( CREATE_OUSTRING( "SUM" ), OUString( OOX_DUMP_BASECLASS ), 1 );
        break;
        case BIFF_TOK_ATTR_ASSIGN:
            dumpUnused( bBiff2 ? 1 : 2 );
        break;
        case BIFF_TOK_ATTR_SPACE:
        case BIFF_TOK_ATTR_SPACE | BIFF_TOK_ATTR_VOLATILE:
            switch( getBiff() )
            {
                case BIFF2:
                    bValid = false;
                break;
                case BIFF3:
                    dumpDec< sal_uInt16 >( "leading-spaces" );
                break;
                case BIFF4:
                case BIFF5:
                case BIFF8:
                    dumpDec< sal_uInt8 >( "char-type", mxSpTypes );
                    dumpDec< sal_uInt8 >( "char-count" );
                break;
                case BIFF_UNKNOWN: break;
            }
        break;
        default:
            bValid = false;
    }
    return bValid;
}

bool FormulaObject::dumpNlrToken()
{
    const OUString aRefClass = cfg().getName( mxClasses, BIFF_TOKCLASS_REF );
    const OUString aValClass = cfg().getName( mxClasses, BIFF_TOKCLASS_VAL );

    bool bValid = true;
    sal_uInt8 nType = dumpHex< sal_uInt8 >( "type", mxNlrTypes );
    switch( nType )
    {
        case BIFF_TOK_NLR_ERR:      dumpNlrErrToken();                      break;
        case BIFF_TOK_NLR_ROWR:     dumpNlrColRowToken( aRefClass, false ); break;
        case BIFF_TOK_NLR_COLR:     dumpNlrColRowToken( aRefClass, false ); break;
        case BIFF_TOK_NLR_ROWV:     dumpNlrColRowToken( aValClass, false ); break;
        case BIFF_TOK_NLR_COLV:     dumpNlrColRowToken( aValClass, false ); break;
        case BIFF_TOK_NLR_RANGE:    dumpNlrRangeToken( aRefClass, false );  break;
        case BIFF_TOK_NLR_SRANGE:   dumpNlrRangeToken( aRefClass, true );   break;
        case BIFF_TOK_NLR_SROWR:    dumpNlrColRowToken( aRefClass, true );  break;
        case BIFF_TOK_NLR_SCOLR:    dumpNlrColRowToken( aRefClass, true );  break;
        case BIFF_TOK_NLR_SROWV:    dumpNlrColRowToken( aValClass, true );  break;
        case BIFF_TOK_NLR_SCOLV:    dumpNlrColRowToken( aValClass, true );  break;
        case BIFF_TOK_NLR_RANGEERR: dumpNlrRangeErrToken();                 break;
        default:                    bValid = false;
    }
    return bValid;
}

void FormulaObject::dumpNlrErrToken()
{
    dumpDec< sal_uInt32 >( "delname-idx" );
    mxStack->pushOperand( lclCreateNlr( getErrorName( BIFF_ERR_NAME ) ) );
}

void FormulaObject::dumpNlrColRowToken( const OUString& rTokClass, bool bAddData )
{
    if( bAddData )
    {
        dumpUnused( 4 );
        mxStack->pushOperand( createPlaceHolder(), rTokClass );
        maAddData.push_back( ADDDATA_NLR );
    }
    else
    {
        TokenAddress aPos = dumpTokenAddress( false );
        writeInfoItem( "addr", lclCreateNlr( aPos ) );
        mxStack->pushOperand( out().getLastItemValue(), rTokClass );
    }
}

void FormulaObject::dumpNlrRangeToken( const OUString& rTokClass, bool bAddData )
{
    if( bAddData )
    {
        dumpUnused( 4 );
        mxStack->pushOperand( createPlaceHolder(), rTokClass );
        maAddData.push_back( ADDDATA_NLR );
    }
    else
    {
        TokenAddress aPos = dumpTokenAddress( false );
        writeInfoItem( "addr", lclCreateNlr( aPos ) );
        mxStack->pushOperand( out().getLastItemValue(), rTokClass );
    }
    dumpUnknown( 1 );
    dumpRange( "target-range" );
}

void FormulaObject::dumpNlrRangeErrToken()
{
    dumpDec< sal_uInt32 >( "delname-idx" );
    dumpUnused( 9 );
    mxStack->pushOperand( lclCreateNlr( getErrorName( BIFF_ERR_NAME ) ) );
}

void FormulaObject::dumpAddTokenData()
{
    Output& rOut = out();
    rOut.resetItemIndex();
    for( AddDataTypeVec::const_iterator aIt = maAddData.begin(), aEnd = maAddData.end(); aIt != aEnd; ++aIt )
    {
        AddDataType eType = *aIt;

        {
            ItemGuard aItem( rOut, "#add-data" );
            switch( eType )
            {
                case ADDDATA_NLR:       rOut.writeAscii( "tNlr" );      break;
                case ADDDATA_ARRAY:     rOut.writeAscii( "tArray" );    break;
                case ADDDATA_MEMAREA:   rOut.writeAscii( "tMemArea" );  break;
            }
        }

        size_t nIdx = aIt - maAddData.begin();
        IndentGuard aIndGuard( rOut );
        switch( eType )
        {
            case ADDDATA_NLR:       dumpAddDataNlr( nIdx );     break;
            case ADDDATA_ARRAY:     dumpAddDataArray( nIdx );   break;
            case ADDDATA_MEMAREA:   dumpAddDataMemArea( nIdx ); break;
        }
    }
}

void FormulaObject::dumpAddDataNlr( size_t nIdx )
{
    sal_uInt32 nFlags = dumpHex< sal_uInt32 >( "flags", "NLRADDFLAGS" );
    sal_uInt32 nCount = nFlags & BIFF_TOK_NLR_ADDMASK;
    OUStringBuffer aBuffer;
    for( sal_uInt32 nPos = 0; nPos < nCount; ++nPos )
    {
        Address aPos;
        aPos.read( getBiffStream() );
        OUStringBuffer aAddr;
        StringHelper::appendAddress( aAddr, aPos );
        StringHelper::appendToken( aBuffer, aAddr.makeStringAndClear(), OOX_DUMP_LISTSEP );
    }
    OUString aAddrList = aBuffer.makeStringAndClear();
    writeInfoItem( "stacked-positions", aAddrList );
    mxStack->replaceOnTop( createPlaceHolder( nIdx ), lclCreateNlr( aAddrList ) );
}

void FormulaObject::dumpAddDataArray( size_t nIdx )
{
    sal_uInt32 nCols, nRows;
    dumpConstArrayHeader( nCols, nRows );

    OUStringBuffer aOp;
    TableGuard aTabGuard( out(), 17 );
    for( sal_uInt32 nRow = 0; nRow < nRows; ++nRow )
    {
        OUStringBuffer aArrayLine;
        for( sal_uInt32 nCol = 0; nCol < nCols; ++nCol )
            StringHelper::appendToken( aArrayLine, dumpConstValue(), OOX_DUMP_LISTSEP );
        StringHelper::appendToken( aOp, aArrayLine.makeStringAndClear(), OOX_DUMP_ARRAYSEP );
    }
    StringHelper::enclose( aOp, '{', '}' );
    mxStack->replaceOnTop( createPlaceHolder( nIdx ), aOp.makeStringAndClear() );
}

void FormulaObject::dumpAddDataMemArea( size_t /*nIdx*/ )
{
    dumpRangeList( 0, getBiff() == BIFF8 );
}

// ============================================================================
// ============================================================================

RecordStreamObject::RecordStreamObject( const ObjectBase& rParent, BinaryInputStreamRef xStrm, BiffType eBiff )
{
    construct( rParent, xStrm, eBiff );
}

RecordStreamObject::RecordStreamObject( const OleStorageObject& rParentStrg, const OUString& rStrmName, BiffType eBiff )
{
    construct( rParentStrg, rStrmName, eBiff );
}

RecordStreamObject::~RecordStreamObject()
{
}

void RecordStreamObject::construct( const ObjectBase& rParent, BinaryInputStreamRef xStrm, BiffType eBiff )
{
    BiffObjectBase::construct( rParent, xStrm, eBiff );
    constructRecStrmObj();
}

void RecordStreamObject::construct( const OleStorageObject& rParentStrg, const OUString& rStrmName, BiffType eBiff )
{
    BiffObjectBase::construct( rParentStrg, rStrmName, eBiff );
    constructRecStrmObj();
}

bool RecordStreamObject::implIsValid() const
{
    return isValid( mxHdrObj ) && isValid( mxFmlaObj ) && isValid( mxDffObj ) && BiffObjectBase::implIsValid();
}

void RecordStreamObject::implDumpBody()
{
    BiffInputStream& rStrm = getBiffStream();

    while( rStrm.startNextRecord() )
    {
        // record header
        rStrm.resetRecord( mxHdrObj->isMergeContRec() );
        implPreProcessRecord();
        mxHdrObj->dumpRecordHeader( rStrm );

        // record contents
        if( mxHdrObj->isShowRecBody() )
        {
            IndentGuard aIndGuard( out() );
            if( mxHdrObj->hasRecName( rStrm.getRecId() ) )
                dumpRecordBody();
            else
                dumpRawBinary( rStrm.getRecSize(), false );
        }
        if( !rStrm.isValid() )
            writeInfoItem( "stream-state", OOX_DUMP_ERR_STREAM );

        // postprocessing
        implPostProcessRecord();
        out().emptyLine();
    }
}

void RecordStreamObject::implPreProcessRecord()
{
}

void RecordStreamObject::implDumpRecord()
{
}

void RecordStreamObject::implPostProcessRecord()
{
}

sal_uInt16 RecordStreamObject::dumpRepeatedRecId()
{
    return dumpHex< sal_uInt16 >( "repeated-rec-id", mxHdrObj->getRecNames() );
}

void RecordStreamObject::constructRecStrmObj()
{
    if( BiffObjectBase::implIsValid() )
    {
        mxHdrObj.reset( new RecordHeaderObject( *this ) );
        mxFmlaObj.reset( new FormulaObject( *this ) );
        mxDffObj.reset( new DffDumpObject( *this ) );
    }

    if( isValid() )
        mxSimpleRecs = cfg().getNameList( "SIMPLE-RECORDS" );
}

void RecordStreamObject::dumpRecordBody()
{
    BiffInputStream& rStrm = getBiffStream();
    sal_uInt16 nRecId = rStrm.getRecId();
    if( cfg().hasName( mxSimpleRecs, nRecId ) )
        dumpSimpleRecord( cfg().getName( mxSimpleRecs, nRecId ) );
    else
        implDumpRecord();
    // remaining undumped data
    if( rStrm.getRecPos() == 0 )
        dumpRawBinary( rStrm.getRecSize(), false );
    else
        dumpRemaining( rStrm.getRecLeft() );
}

void RecordStreamObject::dumpSimpleRecord( const OUString& rRecData )
{
    ItemFormat aItemFmt;
    aItemFmt.parse( rRecData );
    dumpItem( aItemFmt );
}

// ============================================================================

WorkbookStreamObject::WorkbookStreamObject( const DumperBase& rDumper )
{
    if( rDumper.isValid() )
    {
        BinaryInputStreamRef xStrm( new BinaryInputStream( rDumper.getRootStorage()->openInputStream( OUString() ) ) );
        RecordStreamObject::construct( rDumper, xStrm, BIFF_UNKNOWN );
        constructWbStrmObj();
    }
}

WorkbookStreamObject::WorkbookStreamObject( const OleStorageObject& rParentStrg, const OUString& rStrmName )
{
    RecordStreamObject::construct( rParentStrg, rStrmName, BIFF_UNKNOWN );
    constructWbStrmObj();
}

WorkbookStreamObject::~WorkbookStreamObject()
{
    if( WorkbookStreamObject::implIsValid() )
    {
        Config& rCfg = cfg();
        rCfg.eraseNameList( "FONTNAMES" );
        rCfg.eraseNameList( "FORMATS" );
    }
}

void WorkbookStreamObject::implPreProcessRecord()
{
    BiffInputStream& rStrm = getBiffStream();
    sal_uInt16 nRecId = rStrm.getRecId();

    // record specific settings
    switch( nRecId )
    {
        case BIFF_ID_CHEND:
            out().decIndent();
        break;
    }

    // special CONTINUE handling
    if( getRecordHeader().isMergeContRec() )
    {
        switch( nRecId )
        {
            case BIFF_ID_OBJ:
            case BIFF_ID_TXO:
            case BIFF_ID_EOF:
            case BIFF_ID_CONT:
                rStrm.resetRecord( false );
            break;
            case BIFF_ID_MSODRAWINGGROUP:
            case BIFF_ID_CHESCHERFORMAT:
                rStrm.resetRecord( true, nRecId );
            break;
        }
    }
}

void WorkbookStreamObject::implDumpRecord()
{
    BiffInputStream& rStrm = getBiffStream();
    sal_uInt16 nRecId = rStrm.getRecId();
    sal_Size nRecSize = rStrm.getRecSize();
    BiffType eBiff = getBiff();

    switch( nRecId )
    {
        case BIFF2_ID_ARRAY:
        case BIFF3_ID_ARRAY:
            dumpRange( "array-range", false );
            dumpHex< sal_uInt16, sal_uInt8 >( eBiff != BIFF2, "flags", "ARRAY-FLAGS" );
            if( eBiff >= BIFF5 ) dumpUnused( 4 );
            getFormulaDumper().dumpCellFormula();
        break;

        case BIFF2_ID_BLANK:
        case BIFF3_ID_BLANK:
            dumpCellHeader( nRecId == BIFF2_ID_BLANK );
        break;

        case BIFF2_ID_BOF:
        case BIFF3_ID_BOF:
        case BIFF4_ID_BOF:
        case BIFF5_ID_BOF:
            dumpHex< sal_uInt16 >( "bof-type", "BOF-BIFFTYPE" );
            dumpHex< sal_uInt16 >( "sheet-type", "BOF-SHEETTYPE" );
            if( nRecSize >= 6 )  dumpDec< sal_uInt16 >( "build-id" );
            if( nRecSize >= 8 )  dumpDec< sal_uInt16 >( "build-year" );
            if( nRecSize >= 12 ) dumpHex< sal_uInt32 >( "history-flags", "BOF-HISTORY-FLAGS" );
            if( nRecSize >= 16 ) dumpDec< sal_uInt32 >( "lowest-ver" );
        break;

        case BIFF2_ID_BOOLERR:
        case BIFF3_ID_BOOLERR:
            dumpCellHeader( nRecId == BIFF2_ID_BOOLERR );
            dumpBoolErr();
        break;

        case BIFF_ID_BOUNDSHEET:
            if( eBiff >= BIFF5 ) dumpHex< sal_uInt32 >( "sheet-stream-pos", "CONV-DEC" );
            if( eBiff >= BIFF5 ) dumpDec< sal_uInt8 >( "visibility", "BOUNDSHEET-VISIBILITY" );
            if( eBiff >= BIFF5 ) dumpDec< sal_uInt8 >( "sheet-type", "BOUNDSHEET-TYPE" );
            dumpString( "sheet-name", BIFF_STR_8BITLENGTH, BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_CH3DDATAFORMAT:
            dumpDec< sal_uInt8 >( "base", "CH3DDATAFORMAT-BASE" );
            dumpDec< sal_uInt8 >( "top", "CH3DDATAFORMAT-TOP" );
        break;

        case BIFF_ID_CHAREAFORMAT:
            dumpRgbColor( "fg-color-rgb" );
            dumpRgbColor( "bg-color-rgb" );
            dumpPatternIdx();
            dumpHex< sal_uInt16 >( "flags", "CHAREAFORMAT-FLAGS" );
            if( eBiff == BIFF8 ) dumpColorIdx( "fg-color-idx" );
            if( eBiff == BIFF8 ) dumpColorIdx( "bg-color-idx" );
        break;

        case BIFF_ID_CHAXESSET:
            dumpDec< sal_uInt16 >( "axesset-id", "CHAXESSET-ID" );
            dumpRect< sal_Int32 >( "position", (eBiff <= BIFF4) ? "CONV-TWIP-TO-CM" : "" );
        break;

        case BIFF_ID_CHAXIS:
            dumpDec< sal_uInt16 >( "axis-type", "CHAXIS-TYPE" );
            if( eBiff <= BIFF4 )
                dumpRect< sal_Int32 >( "position", "CONV-TWIP-TO-CM" );
            else
                dumpUnused( 16 );
        break;

        case BIFF_ID_CHBAR:
            dumpDec< sal_Int16 >( "overlap", "CONV-PERCENT-NEG" );
            dumpDec< sal_Int16 >( "gap", "CONV-PERCENT" );
            dumpHex< sal_uInt16 >( "flags", "CHBAR-FLAGS" );
        break;

        case BIFF_ID_CHCHART:
            dumpRect< sal_Int32 >( "chart-frame", "CONV-PT1616-TO-CM", FORMATTYPE_FIX );
        break;

        case BIFF_ID_CHCHART3D:
            dumpDec< sal_uInt16 >( "rotation-angle", "CONV-DEG" );
            dumpDec< sal_Int16 >( "elevation-angle", "CONV-DEG" );
            dumpDec< sal_uInt16 >( "eye-distance" );
            dumpDec< sal_uInt16 >( "relative-height", "CONV-PERCENT" );
            dumpDec< sal_uInt16 >( "relative-depth", "CONV-PERCENT" );
            dumpDec< sal_uInt16 >( "depth-gap", "CONV-PERCENT" );
            dumpHex< sal_uInt16 >( "flags", "CHCHART3D-FLAGS" );
        break;

        case BIFF_ID_CHTYPEGROUP:
            dumpUnused( 16 );
            dumpHex< sal_uInt16 >( "flags", "CHTYPEGROUP-FLAGS" );
            if( eBiff >= BIFF5 ) dumpDec< sal_uInt16 >( "group-idx" );
        break;

        case BIFF_ID_CHDATAFORMAT:
            dumpDec< sal_Int16 >( "point-idx", "CHDATAFORMAT-POINTIDX" );
            dumpDec< sal_Int16 >( "series-idx" );
            if( eBiff >= BIFF5 ) dumpDec< sal_Int16 >( "format-idx", "CHDATAFORMAT-FORMATIDX" );
            if( eBiff >= BIFF5 ) dumpHex< sal_uInt16 >( "flags", "CHDATAFORMAT-FLAGS" );
        break;

        case BIFF_ID_CHESCHERFORMAT:
            getDffDumper().dump();
        break;

        case BIFF_ID_CHFRAME:
            dumpDec< sal_uInt16 >( "format", "CHFRAME-FORMAT" );
            dumpHex< sal_uInt16 >( "flags", "CHFRAME-FLAGS" );
        break;

        case BIFF_ID_CHFRAMEPOS:
            dumpDec< sal_uInt16 >( "object-type", "CHFRAMEPOS-OBJTYPE" );
            dumpDec< sal_uInt16 >( "size-mode", "CHFRAMEPOS-SIZEMODE" );
            dumpRect< sal_Int32 >( "position", (eBiff <= BIFF4) ? "CONV-TWIP-TO-CM" : "" );
        break;

        case BIFF_ID_CHLABELRANGE:
            dumpDec< sal_uInt16 >( "axis-crossing" );
            dumpDec< sal_uInt16 >( "label-frequency" );
            dumpDec< sal_uInt16 >( "tick-frequency" );
            dumpHex< sal_uInt16 >( "flags", "CHLABELRANGE-FLAGS" );
        break;

        case BIFF_ID_CHLABELRANGE2:
            dumpDec< sal_uInt16 >( "minimum-categ" );
            dumpDec< sal_uInt16 >( "maximum-categ" );
            dumpDec< sal_uInt16 >( "major-unit-value" );
            dumpDec< sal_uInt16 >( "major-unit" );
            dumpDec< sal_uInt16 >( "minor-unit-value" );
            dumpDec< sal_uInt16 >( "minor-unit" );
            dumpDec< sal_uInt16 >( "base-unit" );
            dumpDec< sal_uInt16 >( "axis-crossing-date" );
            dumpHex< sal_uInt16 >( "flags", "CHLABELRANGE2-FLAGS" );
        break;

        case BIFF_ID_CHLEGEND:
            dumpRect< sal_Int32 >( "position", (eBiff <= BIFF4) ? "CONV-TWIP-TO-CM" : "" );
            dumpDec< sal_uInt8 >( "docked-pos", "CHLEGEND-DOCKPOS" );
            dumpDec< sal_uInt8 >( "spacing", "CHLEGEND-SPACING" );
            dumpHex< sal_uInt16 >( "flags", "CHLEGEND-FLAGS" );
        break;

        case BIFF_ID_CHLINEFORMAT:
            dumpRgbColor();
            dumpDec< sal_uInt16 >( "line-type", "CHLINEFORMAT-LINETYPE" );
            dumpDec< sal_Int16 >( "line-weight", "CHLINEFORMAT-LINEWEIGHT" );
            dumpHex< sal_uInt16 >( "flags", "CHLINEFORMAT-FLAGS" );
            if( eBiff == BIFF8 ) dumpColorIdx();
        break;

        case BIFF_ID_CHMARKERFORMAT:
            dumpRgbColor( "border-color-rgb" );
            dumpRgbColor( "fill-color-rgb" );
            dumpDec< sal_uInt16 >( "marker-type", "CHMARKERFORMAT-TYPE" );
            dumpHex< sal_uInt16 >( "flags", "CHMARKERFORMAT-FLAGS" );
            if( eBiff == BIFF8 ) dumpColorIdx( "border-color-idx" );
            if( eBiff == BIFF8 ) dumpColorIdx( "fill-color-idx" );
            if( eBiff == BIFF8 ) dumpDec< sal_Int32 >( "marker-size", "CONV-TWIP-TO-PT" );
        break;

        case BIFF_ID_CHOBJECTLINK:
            dumpDec< sal_uInt16 >( "link-target", "CHOBJECTLINK-TARGET" );
            dumpDec< sal_Int16 >( "series-idx" );
            dumpDec< sal_Int16 >( "point-idx", "CHOBJECTLINK-POINT" );
        break;

        case BIFF_ID_CHPICFORMAT:
            dumpDec< sal_uInt16 >( "bitmap-mode", "CHPICFORMAT-BITMAP-MODE" );
            dumpDec< sal_uInt16 >( "image-format", "CHPICFORMAT-IMAGE-FORMAT" );
            dumpHex< sal_uInt16 >( "flags", "CHPICFORMAT-FLAGS" );
            dumpDec< double >( "scaling-factor" );
        break;

        case BIFF_ID_CHPIE:
            dumpDec< sal_uInt16 >( "angle", "CONV-DEG" );
            if( eBiff >= BIFF5 ) dumpDec< sal_uInt16 >( "hole-size" );
            if( eBiff >= BIFF8 ) dumpHex< sal_uInt16 >( "flags", "CHPIE-FLAGS" );
        break;

        case BIFF_ID_CHPLOTGROWTH:
            dumpFix< sal_Int32 >( "horizontal-growth" );
            dumpFix< sal_Int32 >( "vertical-growth" );
        break;

        case BIFF_ID_CHPROPERTIES:
            dumpHex< sal_uInt16 >( "flags", "CHPROPERTIES-FLAGS" );
            dumpDec< sal_uInt8 >( "empty-cells", "CHPROPERTIES-EMPTYCELLS" );
        break;

        case BIFF_ID_CHSCATTER:
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "bubble-size", "CONV-PERCENT" );
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "size-type", "CHSCATTER-SIZETYPE" );
            if( eBiff == BIFF8 ) dumpHex< sal_uInt16 >( "flags", "CHSCATTER-FLAGS" );
        break;

        case BIFF_ID_CHSERERRORBAR:
            dumpDec< sal_uInt8 >( "type", "CHSERERRORBAR-TYPE" );
            dumpDec< sal_uInt8 >( "source", "CHSERERRORBAR-SOURCE" );
            dumpBool< sal_uInt8 >( "draw-t-shape" );
            dumpBool< sal_uInt8 >( "draw-line" );
            dumpDec< double >( "value" );
            dumpDec< sal_uInt16 >( "custom-count" );
        break;

        case BIFF_ID_CHSERIES:
            dumpDec< sal_uInt16 >( "categories-type", "CHSERIES-TYPE" );
            dumpDec< sal_uInt16 >( "values-type", "CHSERIES-TYPE" );
            dumpDec< sal_uInt16 >( "categories-count" );
            dumpDec< sal_uInt16 >( "values-count" );
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "bubbles-type", "CHSERIES-TYPE" );
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "bubbles-count" );
        break;

        case BIFF_ID_CHSERTRENDLINE:
            switch( dumpDec< sal_uInt8 >( "type", "CHSERTRENDLINE-TYPE" ) )
            {
                case 0:     dumpDec< sal_uInt8 >( "order" );            break;
                case 4:     dumpDec< sal_uInt8 >( "average-period" );   break;
                default:    dumpUnused( 1 );
            }
            dumpDec< double >( "intercept" );
            dumpBool< sal_uInt8 >( "show-equation" );
            dumpBool< sal_uInt8 >( "show-r-sqrare" );
            dumpDec< double >( "forecast-forward" );
            dumpDec< double >( "forecast-backward" );
        break;

        case BIFF_ID_CHSOURCELINK:
            dumpDec< sal_uInt8 >( "link-target", "CHSOURCELINK-TARGET" );
            dumpDec< sal_uInt8 >( "link-type", "CHSOURCELINK-TYPE" );
            dumpHex< sal_uInt16 >( "flags", "CHSOURCELINK-FLAGS" );
            dumpFormatIdx();
            getFormulaDumper().dumpNameFormula();
        break;

        case BIFF_ID_CHSTRING:
            dumpDec< sal_uInt16 >( "text-type", "CHSTRING-TYPE" );
            dumpString( "text", BIFF_STR_8BITLENGTH, BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_CHTEXT:
            dumpDec< sal_uInt8 >( "horizontal-align", "CHTEXT-HORALIGN" );
            dumpDec< sal_uInt8 >( "vertical-align", "CHTEXT-VERALIGN" );
            dumpDec< sal_uInt16 >( "fill-mode", "CHTEXT-FILLMODE" );
            dumpRgbColor();
            dumpRect< sal_Int32 >( "position", (eBiff <= BIFF4) ? "CONV-TWIP-TO-CM" : "" );
            dumpHex< sal_uInt16 >( "flags", "CHTEXT-FLAGS" );
            if( eBiff == BIFF8 ) dumpColorIdx();
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "placement", "CHTEXT-PLACEMENT" );
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "rotation", "TEXTROTATION" );
        break;

        case BIFF_ID_CHTICK:
            dumpDec< sal_uInt8 >( "major-ticks", "CHTICK-TYPE" );
            dumpDec< sal_uInt8 >( "minor-ticks", "CHTICK-TYPE" );
            dumpDec< sal_uInt8 >( "label-position", "CHTICK-LABELPOS" );
            dumpDec< sal_uInt8 >( "fill-mode", "CHTEXT-FILLMODE" );
            dumpRgbColor( "label-color-rgb" );
            dumpUnused( 16 );
            dumpHex< sal_uInt16 >( "flags", "CHTICK-FLAGS" );
            if( eBiff == BIFF8 ) dumpColorIdx( "label-color-idx" );
            if( eBiff == BIFF8 ) dumpDec< sal_uInt16 >( "label-rotation", "TEXTROTATION" );
        break;

        case BIFF_ID_CHUNITPROPERTIES:
            dumpRepeatedRecId();
            dumpUnused( 2 );
            dumpDec< sal_Int16 >( "preset", "CHUNITPROPERTIES-PRESET" );
            dumpDec< double >( "unit" );
            dumpHex< sal_uInt16 >( "flags", "CHUNITPROPERTIES-FLAGS" );
        break;

        case BIFF_ID_CHVALUERANGE:
            dumpDec< double >( "minimum" );
            dumpDec< double >( "maximum" );
            dumpDec< double >( "major-inc" );
            dumpDec< double >( "minor-inc" );
            dumpDec< double >( "axis-crossing" );
            dumpHex< sal_uInt16 >( "flags", "CHVALUERANGE-FLAGS" );
        break;

        case BIFF_ID_CHWRAPPEDRECORD:
            dumpRepeatedRecId();
            dumpUnused( 2 );
        break;

        case BIFF_ID_CODEPAGE:
            getBiffData().setTextEncoding( dumpCodePage() );
            mbHasCodePage = true;
        break;

        case BIFF_ID_COLINFO:
            dumpColRange();
            dumpDec< sal_uInt16 >( "col-width", "CONV-COLWIDTH" );
            dumpXfIdx( "default-xf-idx" );
            dumpHex< sal_uInt16 >( "flags", "COLINFO-FLAGS" );
            dumpUnused( 2 );
        break;

        case BIFF_ID_COLUMNDEFAULT:
            out().resetItemIndex();
            for( sal_Int32 nCol = 0, nCount = dumpColRange(); nCol < nCount; ++nCol )
                dumpXfIdx( "#xf-idx", true );
            dumpUnused( 2 );
        break;

        case BIFF_ID_COLWIDTH:
            dumpColRange( 0, false );
            dumpDec< sal_uInt16 >( "col-width", "CONV-COLWIDTH" );
        break;

        case BIFF_ID_CRN:
        {
            sal_Int32 nCol2 = dumpColIndex( "last-col-idx", false );
            sal_Int32 nCol1 = dumpColIndex( "first-col-idx", false );
            sal_Int32 nRow = dumpRowIndex( "row-idx" );
            for( Address aPos( nCol1, nRow ); rStrm.isValid() && (aPos.mnCol <= nCol2); ++aPos.mnCol )
            {
                MultiItemsGuard aMultiGuard( out() );
                writeAddressItem( "pos", aPos );
                dumpConstValue();
            }
        }
        break;

        case BIFF_ID_DCONNAME:
            dumpString( "source-name", BIFF_STR_8BITLENGTH );
            dumpString( "source-link", BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_DCONREF:
            dumpRange( "source-range", false );
            dumpString( "source-link", BIFF_STR_8BITLENGTH );
        break;

        case BIFF3_ID_DEFROWHEIGHT:
            dumpHex< sal_uInt16 >( "flags", "DEFROWHEIGHT-FLAGS" );
            dumpDec< sal_uInt16 >( "row-height", "CONV-TWIP-TO-PT" );
        break;

        case BIFF2_ID_DIMENSIONS:
        case BIFF3_ID_DIMENSIONS:
            dumpRange( "used-area", true, (nRecId == BIFF3_ID_DIMENSIONS) && (eBiff == BIFF8) );
            if( nRecId == BIFF3_ID_DIMENSIONS ) dumpUnused( 2 );
        break;

        case BIFF2_ID_EXTERNNAME:
        case BIFF3_ID_EXTERNNAME:
        {
            sal_uInt16 nFlags = (eBiff >= BIFF4) ? dumpHex< sal_uInt16 >( "flags", "EXTERNNAME-FLAGS" ) : 0;
            if( eBiff >= BIFF5 )
            {
                if( getFlag< sal_uInt16 >( nFlags, 0x0010 ) )
                {
                    dumpHex< sal_uInt32 >( "storage-id" );
                }
                else
                {
                    dumpDec< sal_uInt16 >( "externsheet-idx" );
                    dumpUnused( 2 );
                }
            }
            OUString aName = dumpString( "name", BIFF_STR_8BITLENGTH, BIFF_STR_8BITLENGTH );
            if( (eBiff == BIFF4) && (aName.getLength() > 0) && (aName[ 0 ] == 1) && (rStrm.getRecLeft() >= 2) )
                getFormulaDumper().dumpNameFormula();
        }
        break;

        case BIFF_ID_EXTERNSHEET:
            if( eBiff == BIFF8 )
            {
                sal_uInt16 nCount = dumpDec< sal_uInt16 >( "ref-count" );
                TableGuard aTabGuard( out(), 10, 17, 24 );
                out().resetItemIndex();
                for( sal_uInt16 nRefId = 0; rStrm.isValid() && (nRefId < nCount); ++nRefId )
                {
                    MultiItemsGuard aMultiGuard( out() );
                    writeEmptyItem( "#ref" );
                    dumpDec< sal_uInt16 >( "supbook-idx" );
                    dumpDec< sal_Int16 >( "first-sheet", "EXTERNSHEET-IDX" );
                    dumpDec< sal_Int16 >( "last-sheet", "EXTERNSHEET-IDX" );
                }
            }
            else
            {
                dumpByteString( "encoded-url", BIFF_STR_8BITLENGTH );
            }
        break;

        case BIFF2_ID_FONT:
        case BIFF3_ID_FONT:
            dumpFontRec();
        break;

        case BIFF2_ID_FORMAT:
        case BIFF4_ID_FORMAT:
            dumpFormatRec();
        break;

        case BIFF2_ID_FORMULA:
        case BIFF3_ID_FORMULA:
        case BIFF4_ID_FORMULA:
            dumpCellHeader( eBiff == BIFF2 );
            dumpFormulaResult();
            dumpHex< sal_uInt16, sal_uInt8 >( eBiff != BIFF2, "flags", "FORMULA-FLAGS" );
            if( eBiff >= BIFF5 ) dumpUnused( 4 );
            getFormulaDumper().dumpCellFormula();
        break;

        case BIFF_ID_FOOTER:
            if( rStrm.getRecLeft() > 0 )
                dumpString( "footer", BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_HEADER:
            if( rStrm.getRecLeft() > 0 )
                dumpString( "header", BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_HLINK:
            dumpRange();
            if( cfg().getStringOption( dumpGuid( "guid" ), OUString() ).equalsAscii( "StdHlink" ) )
            {
                dumpUnknown( 4 );
                sal_uInt32 nFlags = dumpHex< sal_uInt32 >( "flags", "HLINK-FLAGS" );
                if( getFlag( nFlags, BIFF_HLINK_DISPLAY ) )
                    dumpOleString( "display", true );
                if( getFlag( nFlags, BIFF_HLINK_FRAME ) )
                    dumpOleString( "target-frame", true );
                if( getFlag( nFlags, BIFF_HLINK_TARGET ) )
                {
                    if( getFlag( nFlags, BIFF_HLINK_UNC ) )
                    {
                        dumpOleString( "unc-path", true );
                    }
                    else
                    {
                        OUString aGuid = cfg().getStringOption( dumpGuid( "content-type" ), OUString() );
                        if( aGuid.equalsAscii( "FileMoniker" ) )
                        {
                            dumpDec< sal_Int16 >( "up-level" );
                            dumpOleString( "dos-name", false );
                            dumpUnknown( 24 );
                            if( dumpDec< sal_Int32 >( "total-bytes" ) > 0 )
                            {
                                sal_Int32 nBytes = dumpDec< sal_Int32 >( "filename-bytes" );
                                dumpUnknown( 2 );
                                dumpOleString( "filename", nBytes / 2, true );
                            }
                        }
                        else if( aGuid.equalsAscii( "URLMoniker" ) )
                        {
                            sal_Int32 nBytes = dumpDec< sal_Int32 >( "url-bytes" );
                            dumpOleString( "url", nBytes / 2, true );
                        }
                    }
                }
                if( getFlag( nFlags, BIFF_HLINK_LOC ) )
                    dumpOleString( "location", true );
            }
        break;

        case BIFF2_ID_INTEGER:
            dumpCellHeader( true );
            dumpDec< sal_uInt16 >( "value" );
        break;

        case BIFF2_ID_LABEL:
        case BIFF3_ID_LABEL:
        {
            bool bBiff2 = nRecId == BIFF2_ID_LABEL;
            sal_uInt16 nXfIdx = dumpCellHeader( bBiff2 );
            rtl_TextEncoding eOldTextEnc = getBiffData().getTextEncoding();
            getBiffData().setTextEncoding( getBiffData().getXfEncoding( nXfIdx ) );
            dumpString( "value", bBiff2 ? BIFF_STR_8BITLENGTH : BIFF_STR_DEFAULT );
            getBiffData().setTextEncoding( eOldTextEnc );
        }
        break;

        case BIFF_ID_LABELSST:
            dumpCellHeader();
            dumpDec< sal_uInt32 >( "sst-idx" );
        break;

        case BIFF_ID_MSODRAWING:
        case BIFF_ID_MSODRAWINGGROUP:
        case BIFF_ID_MSODRAWINGSEL:
            getDffDumper().dump();
        break;

        case BIFF_ID_MULTBLANK:
        {
            Address aPos = dumpAddress();
            {
                TableGuard aTabGuard( out(), 12 );
                for( ; rStrm.getRecLeft() >= 4; ++aPos.mnCol )
                {
                    MultiItemsGuard aMultiGuard( out() );
                    writeAddressItem( "pos", aPos );
                    dumpXfIdx();
                }
            }
            dumpColIndex( "last-col-idx" );
        }
        break;

        case BIFF_ID_MULTRK:
        {
            Address aPos = dumpAddress();
            {
                TableGuard aTabGuard( out(), 12, 12 );
                for( ; rStrm.getRecLeft() >= 8; ++aPos.mnCol )
                {
                    MultiItemsGuard aMultiGuard( out() );
                    writeAddressItem( "pos", aPos );
                    dumpXfIdx();
                    dumpRk( "value" );
                }
            }
            dumpColIndex( "last-col-idx" );
        }
        break;

        case BIFF2_ID_NAME:
        case BIFF3_ID_NAME:
        {
            dumpHex< sal_uInt16, sal_uInt8 >( eBiff != BIFF2, "flags", "NAME-FLAGS" );
            if( eBiff == BIFF2 ) dumpDec< sal_uInt8 >( "macro-type", "NAME-MACROTYPE-BIFF2" );
            dumpHex< sal_uInt8 >( "keyboard-shortcut" );
            sal_uInt8 nNameLen = dumpDec< sal_uInt8 >( "name-len" );
            sal_uInt16 nFmlaSize = getFormulaDumper().dumpFormulaSize();
            rtl_TextEncoding eTextEnc = getBiffData().getTextEncoding();
            if( eBiff >= BIFF5 )
            {
                bool bBiff8 = eBiff == BIFF8;
                if( bBiff8 ) dumpUnused( 2 ); else dumpDec< sal_uInt16 >( "externsheet-idx", "NAME-SHEETIDX" );
                dumpDec< sal_uInt16 >( "sheet-idx", "NAME-SHEETIDX" );
                sal_uInt8 nMenuLen = dumpDec< sal_uInt8 >( "menu-text-len" );
                sal_uInt8 nDescrLen = dumpDec< sal_uInt8 >( "description-text-len" );
                sal_uInt8 nHelpLen = dumpDec< sal_uInt8 >( "help-text-len" );
                sal_uInt8 nStatusLen = dumpDec< sal_uInt8 >( "statusbar-text-len" );
                writeStringItem( "name", bBiff8 ? rStrm.readUniString( nNameLen ) : rStrm.readCharArray( nNameLen, eTextEnc ) );
                getFormulaDumper().dumpNameFormula( 0, nFmlaSize );
                if( nMenuLen > 0 ) writeStringItem( "menu-text", bBiff8 ? rStrm.readUniString( nMenuLen ) : rStrm.readCharArray( nMenuLen, eTextEnc ) );
                if( nDescrLen > 0 ) writeStringItem( "description-text", bBiff8 ? rStrm.readUniString( nDescrLen ) : rStrm.readCharArray( nDescrLen, eTextEnc ) );
                if( nHelpLen > 0 ) writeStringItem( "help-text", bBiff8 ? rStrm.readUniString( nHelpLen ) : rStrm.readCharArray( nHelpLen, eTextEnc ) );
                if( nStatusLen > 0 ) writeStringItem( "statusbar-text", bBiff8 ? rStrm.readUniString( nStatusLen ) : rStrm.readCharArray( nStatusLen, eTextEnc ) );
            }
            else
            {
                writeStringItem( "name", rStrm.readCharArray( nNameLen, eTextEnc ) );
                getFormulaDumper().dumpNameFormula( 0, nFmlaSize );
                if( eBiff == BIFF2 ) getFormulaDumper().dumpFormulaSize();
            }
        }
        break;

        case BIFF2_ID_NUMBER:
        case BIFF3_ID_NUMBER:
            dumpCellHeader( nRecId == BIFF2_ID_NUMBER );
            dumpDec< double >( "value" );
        break;

        case BIFF_ID_OBJ:
            dumpObjRec();
        break;

        case BIFF_ID_PANE:
            dumpDec< sal_uInt16 >( "x-pos", "CONV-TWIP-TO-CM" );
            dumpDec< sal_uInt16 >( "y-pos", "CONV-TWIP-TO-CM" );
            dumpAddress( "first-visible-cell" );
            dumpDec< sal_uInt8 >( "active-pane", "PANE-ID" );
        break;

        case BIFF_ID_PROJEXTSHEET:
            dumpDec< sal_uInt8 >( "sheet-type", "PROJEXTSHEET-TYPE" );
            dumpUnused( 1 );
            dumpByteString( "sheet-link", BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_RK:
            dumpCellHeader();
            dumpRk( "value" );
        break;

        case BIFF2_ID_ROW:
        {
            dumpRowIndex();
            dumpColIndex( "first-used-col-idx" );
            dumpColIndex( "first-free-col-idx" );
            dumpHex< sal_uInt16 >( "row-height", "ROW-HEIGHT" );
            dumpUnused( 2 );
            bool bHasDefXf = dumpBool< sal_uInt8 >( "has-default-xf" );
            dumpDec< sal_uInt16 >( "cell-offset" );
            if( bHasDefXf ) dumpXfIdx( "default-format", true );
            if( bHasDefXf ) dumpXfIdx( "default-xf-idx", false );
        }
        break;

        case BIFF3_ID_ROW:
            dumpRowIndex();
            dumpColIndex( "first-used-col-idx" );
            dumpColIndex( "first-free-col-idx" );
            dumpHex< sal_uInt16 >( "row-height", "ROW-HEIGHT" );
            dumpUnused( (eBiff <= BIFF4) ? 2 : 4 );
            if( eBiff <= BIFF4 ) dumpDec< sal_uInt16 >( "cell-offset" );
            dumpHex< sal_uInt32 >( "flags", "ROW-FLAGS" );
        break;

        case BIFF_ID_RSTRING:
        {
            sal_uInt16 nXfIdx = dumpCellHeader();
            rtl_TextEncoding eOldTextEnc = getBiffData().getTextEncoding();
            getBiffData().setTextEncoding( getBiffData().getXfEncoding( nXfIdx ) );
            dumpString( "value" );
            getBiffData().setTextEncoding( eOldTextEnc );
            BiffRichStringFontIdVec aFontIds;
            RichString::importFontIds( aFontIds, rStrm, eBiff == BIFF8 );
            writeRichStringFontIds( aFontIds );
        }
        break;

        case BIFF_ID_SCL:
        {
            sal_uInt16 nNum = dumpDec< sal_uInt16 >( "numerator" );
            sal_uInt16 nDen = dumpDec< sal_uInt16 >( "denominator" );
            if( nDen > 0 ) writeDecItem( "current-zoom", static_cast< sal_uInt16 >( nNum * 100 / nDen ), "CONV-PERCENT" );
        }
        break;

        case BIFF_ID_SCREENTIP:
            dumpRepeatedRecId();
            dumpRange();
            dumpNullString( "tooltip", true );
        break;

        case BIFF_ID_SELECTION:
            dumpDec< sal_uInt8 >( "pane", "PANE-ID" );
            dumpAddress( "active-cell" );
            dumpDec< sal_uInt16 >( "list-idx" );
            dumpRangeList( "selection", false );
        break;

        case BIFF_ID_SETUP:
            dumpDec< sal_uInt16 >( "paper-size", "SETUP-PAPERSIZE" );
            dumpDec< sal_uInt16 >( "scaling", "CONV-PERCENT" );
            dumpDec< sal_uInt16 >( "start-page" );
            dumpDec< sal_uInt16 >( "scale-to-width", "SETUP-SCALETOPAGES" );
            dumpDec< sal_uInt16 >( "scale-to-height", "SETUP-SCALETOPAGES" );
            dumpHex< sal_uInt16 >( "flags", "SETUP-FLAGS" );
            if( eBiff >= BIFF5 )
            {
                dumpDec< sal_uInt16 >( "horizontal-res", "SETUP-DPI" );
                dumpDec< sal_uInt16 >( "vertical-res", "SETUP-DPI" );
                dumpDec< double >( "header-margin", "CONV-INCH-TO-CM" );
                dumpDec< double >( "footer-margin", "CONV-INCH-TO-CM" );
                dumpDec< sal_uInt16 >( "copies" );
            }
        break;

        case BIFF_ID_SHEETHEADER:
            dumpHex< sal_uInt32 >( "substream-size", "CONV-DEC" );
            dumpByteString( "sheet-name", BIFF_STR_8BITLENGTH );
        break;

        case BIFF_ID_SHRFMLA:
            dumpRange( "formula-range", false );
            dumpUnused( 1 );
            dumpDec< sal_uInt8 >( "cell-count" );
            getFormulaDumper().dumpCellFormula();
        break;

        case BIFF_ID_SST:
            dumpDec< sal_uInt32 >( "string-cell-count" );
            dumpDec< sal_uInt32 >( "sst-size" );
            out().resetItemIndex();
            while( rStrm.isValid() && (rStrm.getRecLeft() >= 3) )
                dumpUniString( "#entry" );
        break;

        case BIFF2_ID_STRING:
        case BIFF3_ID_STRING:
            dumpString( "result", ((nRecId == BIFF2_ID_STRING) && (eBiff <= BIFF4)) ? BIFF_STR_8BITLENGTH : BIFF_STR_DEFAULT );
        break;

        case BIFF_ID_STYLE:
        {
            sal_uInt16 nFlags = dumpHex< sal_uInt16 >( "flags", "STYLE-FLAGS" );
            if( getFlag( nFlags, BIFF_STYLE_BUILTIN ) )
            {
                dumpDec< sal_uInt8 >( "style-idx", "STYLE-BUILTIN" );
                dumpDec< sal_uInt8 >( "outline-level" );
            }
            else
                dumpString( "style-name", BIFF_STR_8BITLENGTH );
        }
        break;

        case BIFF_ID_SUPBOOK:
        {
            sal_uInt16 nCount = dumpDec< sal_uInt16 >( "sheet-count" );
            if( rStrm.getRecLeft() == 2 )
                dumpHex< sal_uInt16 >( "special-key", "SUPBOOK-KEY" );
            else
            {
                dumpString( "workbook-url" );
                out().resetItemIndex();
                for( sal_uInt16 nSheet = 0; rStrm.isValid() && (nSheet < nCount); ++nSheet )
                    dumpString( "#sheet-name" );
            }
        }
        break;

        case BIFF_ID_SXDI:
        {
            dumpDec< sal_uInt16 >( "field-idx" );
            dumpDec< sal_uInt16 >( "function", "SXDI-FUNC" );
            dumpDec< sal_uInt16 >( "data-format", "SXDI-FORMAT" );
            dumpDec< sal_uInt16 >( "format-basefield-idx" );
            dumpDec< sal_uInt16 >( "format-baseitem-idx", "SXDI-BASEITEM" );
            dumpFormatIdx();
            sal_uInt16 nNameLen = dumpDec< sal_uInt16 >( "item-name-len", "SX-NAMELEN" );
            if( nNameLen != BIFF_PT_NOSTRING )
                writeStringItem( "item-name", rStrm.readUniString( nNameLen ) );
        }
        break;

        case BIFF_ID_SXEXT:
            if( eBiff == BIFF8 )
            {
                dumpHex< sal_uInt16 >( "flags", "SXEXT-FLAGS" );
                dumpDec< sal_uInt16 >( "param-string-count" );
                dumpDec< sal_uInt16 >( "sql-statement-string-count" );
                dumpDec< sal_uInt16 >( "webquery-postmethod-string-count" );
                dumpDec< sal_uInt16 >( "server-pagefields-string-count" );
                dumpDec< sal_uInt16 >( "odbc-connection-string-count" );
            }
        break;

        case BIFF_ID_SXIVD:
            out().resetItemIndex();
            for( sal_Size nIdx = 0, nCount = rStrm.getRecLeft() / 2; nIdx < nCount; ++nIdx )
                dumpDec< sal_uInt16 >( "#field-idx" );
        break;

        case BIFF_ID_SXLI:
            if( mnPTSxliIdx < 2 )
            {
                sal_uInt16 nCount = (mnPTSxliIdx == 0) ? mnPTRowFields : mnPTColFields;
                sal_Size nLineSize = 8 + 2 * nCount;
                out().resetItemIndex();
                while( rStrm.getRecLeft() >= nLineSize )
                {
                    writeEmptyItem( "#line-data" );
                    IndentGuard aIndGuard( out() );
                    MultiItemsGuard aMultiGuard( out() );
                    dumpDec< sal_uInt16 >( "ident-count" );
                    dumpDec< sal_uInt16 >( "item-type", "SXLI-ITEMTYPE" );
                    dumpDec< sal_uInt16 >( "used-count" );
                    dumpHex< sal_uInt16 >( "flags", "SXLI-FLAGS" );
                    OUStringBuffer aItemList;
                    for( sal_uInt16 nIdx = 0; nIdx < nCount; ++nIdx )
                        StringHelper::appendToken( aItemList, in().readValue< sal_uInt16 >() );
                    writeInfoItem( "item-idxs", aItemList.makeStringAndClear() );
                }
                ++mnPTSxliIdx;
            }
        break;

        case BIFF_ID_SXSTRING:
            dumpString( "value" );
        break;

        case BIFF_ID_SXVD:
        {
            dumpDec< sal_uInt16 >( "axis-type", "SXVD-AXISTYPE" );
            dumpDec< sal_uInt16 >( "subtotal-count" );
            dumpHex< sal_uInt16 >( "subtotals", "SXVD-SUBTOTALS" );
            dumpDec< sal_uInt16 >( "item-count" );
            sal_uInt16 nNameLen = dumpDec< sal_uInt16 >( "field-name-len", "SX-NAMELEN" );
            if( nNameLen != BIFF_PT_NOSTRING )
                writeStringItem( "field-name", rStrm.readUniString( nNameLen ) );
        }
        break;

        case BIFF_ID_SXVDEX:
            dumpHex< sal_uInt32 >( "flags", "SXVDEX-FLAGS" );
            dumpDec< sal_uInt16 >( "autosort-basefield-idx" );
            dumpDec< sal_uInt16 >( "autoshow-basefield-idx" );
            dumpFormatIdx();
        break;

        case BIFF_ID_SXVI:
        {
            dumpDec< sal_uInt16 >( "item-type", "SXVI-ITEMTYPE" );
            dumpHex< sal_uInt16 >( "flags", "SXVI-FLAGS" );
            dumpDec< sal_uInt16 >( "cache-idx" );
            sal_uInt16 nNameLen = dumpDec< sal_uInt16 >( "item-name-len", "SX-NAMELEN" );
            if( nNameLen != BIFF_PT_NOSTRING )
                writeStringItem( "item-name", rStrm.readUniString( nNameLen ) );
        }
        break;

        case BIFF_ID_SXVIEW:
        {
            dumpRange( "output-range" );
            dumpRowIndex( "first-header-row-idx" );
            dumpAddress( "first-data-pos" );
            dumpDec< sal_uInt16 >( "cache-idx" );
            dumpUnused( 2 );
            dumpDec< sal_uInt16 >( "default-data-axis", "SXVD-AXISTYPE" );
            dumpDec< sal_Int16 >( "default-data-pos" );
            dumpDec< sal_uInt16 >( "field-count" );
            mnPTRowFields = dumpDec< sal_uInt16 >( "row-field-count" );
            mnPTColFields = dumpDec< sal_uInt16 >( "column-field-count" );
            dumpDec< sal_uInt16 >( "page-field-count" );
            dumpDec< sal_uInt16 >( "data-field-count" );
            dumpDec< sal_uInt16 >( "data-row-count" );
            dumpDec< sal_uInt16 >( "data-column-count" );
            dumpHex< sal_uInt16 >( "flags", "SXVIEW-FLAGS" );
            dumpDec< sal_uInt16 >( "auto-format-idx" );
            sal_uInt16 nTabNameLen = dumpDec< sal_uInt16 >( "table-name-len" );
            sal_uInt16 nDataNameLen = dumpDec< sal_uInt16 >( "data-name-len" );
            writeStringItem( "table-name", rStrm.readUniString( nTabNameLen ) );
            writeStringItem( "data-name", rStrm.readUniString( nDataNameLen ) );
            mnPTSxliIdx = 0;
        }
        break;

        case BIFF_ID_WINDOW1:
            dumpDec< sal_uInt16 >( "window-x", "CONV-TWIP-TO-CM" );
            dumpDec< sal_uInt16 >( "window-y", "CONV-TWIP-TO-CM" );
            dumpDec< sal_uInt16 >( "window-width", "CONV-TWIP-TO-CM" );
            dumpDec< sal_uInt16 >( "window-height", "CONV-TWIP-TO-CM" );
            if( eBiff <= BIFF4 )
            {
                dumpBool< sal_uInt8 >( "hidden" );
            }
            else
            {
                dumpHex< sal_uInt16 >( "flags", "WINDOW1-FLAGS" );
                dumpDec< sal_uInt16 >( "active-tab" );
                dumpDec< sal_uInt16 >( "first-visible-tab" );
                dumpDec< sal_uInt16 >( "selected-tabs" );
                dumpDec< sal_uInt16 >( "tabbar-ratio", "WINDOW1-TABBARRATIO" );
            }
        break;

        case BIFF2_ID_WINDOW2:
            dumpBool< sal_uInt8 >( "show-formulas" );
            dumpBool< sal_uInt8 >( "show-gridlines" );
            dumpBool< sal_uInt8 >( "show-headings" );
            dumpBool< sal_uInt8 >( "frozen-panes" );
            dumpBool< sal_uInt8 >( "show-zeros" );
            dumpAddress( "first-visible-cell" );
            dumpBool< sal_uInt8 >( "auto-grid-color" );
            dumpRgbColor( "grid-color-rgb" );
        break;

        case BIFF3_ID_WINDOW2:
            dumpHex< sal_uInt16 >( "flags", "WINDOW2-FLAGS" );
            dumpAddress( "first-visible-cell" );
            if( eBiff == BIFF8 )
            {
                dumpColorIdx( "grid-color-idx" );
                dumpUnused( 2 );
                if( rStrm.getRecLeft() >= 8 )
                {
                    dumpDec< sal_uInt16 >( "pagebreak-zoom", "CONV-PERCENT" );
                    dumpDec< sal_uInt16 >( "normal-zoom", "CONV-PERCENT" );
                    dumpUnused( 4 );
                }
            }
            else
                dumpRgbColor( "grid-color-rgb" );
        break;

        case BIFF_ID_XCT:
            dumpDec< sal_uInt16 >( "crn-count" );
            if( eBiff == BIFF8 ) dumpDec< sal_Int16 >( "sheet-idx" );
        break;

        case BIFF2_ID_XF:
        case BIFF3_ID_XF:
        case BIFF4_ID_XF:
        case BIFF5_ID_XF:
            dumpXfRec();
        break;
    }
}

void WorkbookStreamObject::implPostProcessRecord()
{
    BiffInputStream& rStrm = getBiffStream();
    sal_uInt16 nRecId = rStrm.getRecId();

    // record specific settings
    switch( nRecId )
    {
        case BIFF_ID_CHBEGIN:
            out().incIndent();
        break;
    }
}

void WorkbookStreamObject::constructWbStrmObj()
{
    if( isValid() )
    {
        Config& rCfg = cfg();
        mxColors        = rCfg.getNameList( "COLORS" );
        mxBorderStyles  = rCfg.getNameList( "BORDERSTYLES" );
        mxFillPatterns  = rCfg.getNameList( "FILLPATTERNS" );
        mxFontNames     = rCfg.createNameList< ConstList >( "FONTNAMES" );
        mxFontNames->setName( 0, createFontName( CREATE_OUSTRING( "Arial" ), 200, false, false ) );
        mxFormats       = rCfg.createNameList< ConstList >( "FORMATS" );
        mxFormats->includeList( rCfg.getNameList( "BUILTIN-FORMATS" ) );
        mnFormatIdx = 0;
        mnPTRowFields = 0;
        mnPTColFields = 0;
        mnPTSxliIdx = 0;
        mbHasCodePage = false;
    }
}

OUString WorkbookStreamObject::createFontName( const OUString& rName, sal_uInt16 nHeight, bool bBold, bool bItalic ) const
{
    OUStringBuffer aName( rName );
    StringHelper::enclose( aName, '\'' );
    StringHelper::appendToken( aName, cfg().getName( "CONV-TWIP-TO-PT", nHeight ), '/' );
    if( bBold )
        StringHelper::appendToken( aName, CREATE_OUSTRING( "bold" ), '/' );
    if( bItalic )
        StringHelper::appendToken( aName, CREATE_OUSTRING( "italic" ), '/' );
    return aName.makeStringAndClear();
}

sal_uInt16 WorkbookStreamObject::dumpPatternIdx( const sal_Char* pcName, bool b16Bit )
{
    return dumpDec< sal_uInt16, sal_uInt8 >( b16Bit, pcName ? pcName : "fill-pattern", mxFillPatterns );
}

sal_uInt16 WorkbookStreamObject::dumpColorIdx( const sal_Char* pcName, bool b16Bit )
{
    return dumpDec< sal_uInt16, sal_uInt8 >( b16Bit, pcName ? pcName : "color-idx", mxColors );
}

sal_uInt16 WorkbookStreamObject::dumpFontIdx( const sal_Char* pcName, bool b16Bit )
{
    return dumpDec< sal_uInt16, sal_uInt8 >( b16Bit, pcName ? pcName : "font-idx", "FONTNAMES" );
}

sal_uInt16 WorkbookStreamObject::dumpFormatIdx( const sal_Char* pcName )
{
    return dumpDec< sal_uInt16, sal_uInt8 >( getBiff() >= BIFF5, pcName ? pcName : "fmt-idx", "FORMATS" );
}

sal_uInt16 WorkbookStreamObject::dumpXfIdx( const sal_Char* pcName, bool bBiff2Style )
{
    if( !pcName ) pcName = "xf-idx";
    sal_uInt16 nXfIdx = 0;
    if( bBiff2Style )
    {
        dumpHex< sal_uInt8 >( pcName, "CELL-XFINDEX" );
        dumpHex< sal_uInt8 >( "fmt-font-idx", "CELL-XFFORMAT" );
        dumpHex< sal_uInt8 >( "style", "CELL-XFSTYLE" );
    }
    else
        nXfIdx = dumpDec< sal_uInt16 >( pcName );
    return nXfIdx;
}

sal_uInt16 WorkbookStreamObject::dumpCellHeader( bool bBiff2Style )
{
    dumpAddress();
    return dumpXfIdx( 0, bBiff2Style );
}

void WorkbookStreamObject::dumpBoolErr()
{
    MultiItemsGuard aMultiGuard( out() );
    sal_uInt8 nValue = dumpHex< sal_uInt8 >( "value" );
    bool bErrCode = dumpBool< sal_uInt8 >( "is-errorcode" );
    if( bErrCode )
        writeErrorCodeItem( "errorcode", nValue );
    else
        writeBooleanItem( "boolean", nValue );
}

void WorkbookStreamObject::dumpFontRec()
{
    sal_uInt16 nFontId = getBiffData().getFontCount();
    out().resetItemIndex( nFontId );
    writeEmptyItem( "#font" );
    sal_uInt16 nHeight = dumpDec< sal_uInt16 >( "height", "CONV-TWIP-TO-PT" );
    sal_uInt16 nFlags = dumpHex< sal_uInt16 >( "flags", "FONT-FLAGS" );
    bool bBold = getFlag( nFlags, BIFF_FONTFLAG_BOLD );
    bool bItalic = getFlag( nFlags, BIFF_FONTFLAG_ITALIC );
    rtl_TextEncoding eFontEnc = RTL_TEXTENCODING_DONTKNOW;
    if( getBiff() >= BIFF3 )
        dumpColorIdx();
    if( getBiff() >= BIFF5 )
    {
        bBold = dumpDec< sal_uInt16 >( "weight", "FONT-WEIGHT" ) > 450;
        dumpDec< sal_uInt16 >( "escapement", "FONT-ESCAPEMENT" );
        dumpDec< sal_uInt8 >( "underline", "FONT-UNDERLINE" );
        dumpDec< sal_uInt8 >( "family", "FONT-FAMILY" );
        sal_uInt8 nCharSet = dumpDec< sal_uInt8 >( "charset", "FONT-CHARSET" );
        eFontEnc = rtl_getTextEncodingFromWindowsCharset( nCharSet );
        dumpUnused( 1 );
    }
    OUString aName = dumpString( "name", BIFF_STR_8BITLENGTH, BIFF_STR_8BITLENGTH );

    // append font data to vector
    mxFontNames->setName( nFontId, createFontName( aName, nHeight, bBold, bItalic ) );

    // store font encoding
    getBiffData().appendFontEncoding( eFontEnc );

    // set font encoding as default text encoding in case of missing CODEPAGE record
    if( !mbHasCodePage && (nFontId == 0) )
        getBiffData().setTextEncoding( eFontEnc );
}

void WorkbookStreamObject::dumpFormatRec()
{
    sal_uInt16 nFormatIdx = 0;
    switch( getBiff() )
    {
        case BIFF2:
        case BIFF3:
            nFormatIdx = mnFormatIdx++;
            out().resetItemIndex( nFormatIdx );
            writeEmptyItem( "#fmt" );
        break;
        case BIFF4:
            nFormatIdx = mnFormatIdx++;
            out().resetItemIndex( nFormatIdx );
            writeEmptyItem( "#fmt" );
            dumpUnused( 2 );
        break;
        case BIFF5:
        case BIFF8:
            getBiffStream() >> nFormatIdx;
            out().resetItemIndex( nFormatIdx );
            writeEmptyItem( "#fmt" );
            writeDecItem( "fmt-idx", nFormatIdx );
        break;
        case BIFF_UNKNOWN: break;
    }
    OUString aFormat = dumpString( "format", BIFF_STR_8BITLENGTH );
    mxFormats->setName( nFormatIdx, aFormat );
}

void WorkbookStreamObject::dumpXfRec()
{
    sal_uInt16 nXfId = getBiffData().getXfCount();
    out().resetItemIndex( nXfId );
    writeEmptyItem( "#xf" );
    sal_uInt16 nFontId = dumpFontIdx( 0, getBiff() >= BIFF5 );
    switch( getBiff() )
    {
        case BIFF2:
            dumpUnused( 1 );
            dumpHex< sal_uInt8 >( "type-flags", "XF-TYPEFLAGS" );
            dumpHex< sal_uInt8 >( "style-flags", "XF-STYLEFLAGS" );
        break;
        case BIFF3:
            dumpFormatIdx();
            dumpHex< sal_uInt8 >( "type-flags", "XF-TYPEFLAGS" );
            dumpHex< sal_uInt8 >( "used-attributes", "XF-USEDATTRIBS-FLAGS" );
            dumpHex< sal_uInt16 >( "alignment", "XF-ALIGNMENT" );
            dumpHex< sal_uInt16 >( "fill-style", "XF-FILL" );
            dumpHex< sal_uInt32 >( "border-style", "XF-BORDER" );
        break;
        case BIFF4:
            dumpFormatIdx();
            dumpHex< sal_uInt16 >( "type-flags", "XF-TYPEFLAGS" );
            dumpHex< sal_uInt8 >( "alignment", "XF-ALIGNMENT" );
            dumpHex< sal_uInt8 >( "used-attributes", "XF-USEDATTRIBS-FLAGS" );
            dumpHex< sal_uInt16 >( "fill-style", "XF-FILL" );
            dumpHex< sal_uInt32 >( "border-style", "XF-BORDER" );
        break;
        case BIFF5:
            dumpFormatIdx();
            dumpHex< sal_uInt16 >( "type-flags", "XF-TYPEFLAGS" );
            dumpHex< sal_uInt8 >( "alignment", "XF-ALIGNMENT" );
            dumpHex< sal_uInt8 >( "orientation", "XF-ORIENTATTRIBS" );
            dumpHex< sal_uInt32 >( "fill-style", "XF-FILL" );
            dumpHex< sal_uInt32 >( "border-style", "XF-BORDER" );
        break;
        case BIFF8:
            dumpFormatIdx();
            dumpHex< sal_uInt16 >( "type-flags", "XF-TYPEFLAGS" );
            dumpHex< sal_uInt8 >( "alignment", "XF-ALIGNMENT" );
            dumpDec< sal_uInt8 >( "rotation", "TEXTROTATION" );
            dumpHex< sal_uInt8 >( "text-flags", "XF-TEXTFLAGS" );
            dumpHex< sal_uInt8 >( "used-attributes", "XF-USEDATTRIBS-FLAGS" );
            dumpHex< sal_uInt16 >( "border-style", "XF-BORDERSTYLE" );
            dumpHex< sal_uInt16 >( "border-color1", "XF-BORDERCOLOR1" );
            dumpHex< sal_uInt32 >( "border-color2", "XF-BORDERCOLOR2" );
            dumpHex< sal_uInt16 >( "fill-color", "XF-FILLCOLOR" );
        break;
        case BIFF_UNKNOWN: break;
    }
    getBiffData().appendXfFontId( nFontId );
}

void WorkbookStreamObject::dumpObjRec()
{
    switch( getBiff() )
    {
        case BIFF5: dumpObjRecBiff5();  break;
        case BIFF8: dumpObjRecBiff8();  break;
        default:;
    }
}

void WorkbookStreamObject::dumpObjRecBiff5()
{
    dumpDec< sal_uInt32 >( "obj-count" );
    sal_uInt16 nObjType = dumpDec< sal_uInt16 >( "obj-type", "OBJ-TYPE" );
    dumpDec< sal_uInt16 >( "obj-id" );
    dumpHex< sal_uInt16 >( "flags", "OBJ-FLAGS" );
    getDffDumper().dumpDffClientRect();
    dumpDec< sal_uInt16 >( "macro-len" );
    dumpUnused( 6 );
    switch( nObjType )
    {
        case BIFF_OBJCMO_LINE:
            dumpColorIdx( "line-color-idx", false );
            dumpDec< sal_uInt8 >( "line-type", "OBJ-LINETYPE" );
            dumpDec< sal_uInt8 >( "line-weight", "OBJ-LINEWEIGHT" );
            dumpHex< sal_uInt8 >( "line-flags", "OBJ-FLAGS-AUTO" );
            dumpHex< sal_uInt16 >( "line-end", "OBJ-LINEENDS" );
            dumpDec< sal_uInt8 >( "line-direction", "OBJ-LINEDIR" );
            dumpUnused( 1 );
        break;
        case BIFF_OBJCMO_RECTANGLE:
            dumpColorIdx( "back-color-idx", false );
            dumpColorIdx( "patt-color-idx", false );
            dumpPatternIdx( 0, false );
            dumpHex< sal_uInt8 >( "area-flags", "OBJ-FLAGS-AUTO" );
            dumpColorIdx( "line-color-idx", false );
            dumpDec< sal_uInt8 >( "line-type", "OBJ-LINETYPE" );
            dumpDec< sal_uInt8 >( "line-weight", "OBJ-LINEWEIGHT" );
            dumpHex< sal_uInt8 >( "line-flags", "OBJ-FLAGS-AUTO" );
            dumpHex< sal_uInt16 >( "frame-style", "OBJ-FRAMESTYLE-FLAGS" );
        break;
        case BIFF_OBJCMO_CHART:
            dumpColorIdx( "back-color-idx", false );
            dumpColorIdx( "patt-color-idx", false );
            dumpPatternIdx( 0, false );
            dumpHex< sal_uInt8 >( "area-flags", "OBJ-FLAGS-AUTO" );
            dumpColorIdx( "line-color-idx", false );
            dumpDec< sal_uInt8 >( "line-type", "OBJ-LINETYPE" );
            dumpDec< sal_uInt8 >( "line-weight", "OBJ-LINEWEIGHT" );
            dumpHex< sal_uInt8 >( "line-flags", "OBJ-FLAGS-AUTO" );
            dumpHex< sal_uInt16 >( "frame-style", "OBJ-FRAMESTYLE-FLAGS" );
            dumpHex< sal_uInt16 >( "chart-flags", "OBJ-CHART-FLAGS" );
            dumpUnused( 16 );
        break;
    }
}

void WorkbookStreamObject::dumpObjRecBiff8()
{
    Output& rOut = out();
    BiffInputStream& rStrm = getBiffStream();
    NameListRef xRecNames = cfg().getNameList( "OBJ-RECNAMES" );
    bool bLinked = false;
    bool bControl = false;
    bool bCtlsStrm = false;
    bool bLoop = true;
    while( bLoop && (rStrm.getRecLeft() >= 4) )
    {
        rOut.emptyLine();
        sal_uInt16 nSubRecId, nSubRecSize;
        {
            MultiItemsGuard aMultiGuard( rOut );
            writeEmptyItem( "OBJREC" );
            writeHexItem( "pos", rStrm.getRecPos() );
            rStrm >> nSubRecId >> nSubRecSize;
            writeHexItem( "size", nSubRecSize );
            writeHexItem( "id", nSubRecId, xRecNames );
        }

        sal_uInt32 nSubRecStart = rStrm.getRecPos();
        // sometimes the last subrecord has an invalid length
        sal_uInt32 nRealRecSize = ::std::min< sal_uInt32 >( nSubRecSize, rStrm.getRecLeft() );
        sal_uInt32 nSubRecEnd = nSubRecStart + nRealRecSize;

        IndentGuard aIndGuard( rOut );
        switch( nSubRecId )
        {
            case BIFF_ID_OBJPIOGRBIT:
            {
                sal_uInt16 nFlags = dumpHex< sal_uInt16 >( "flags", "OBJPIOGRBIT-FLAGS" );
                bLinked = getFlag( nFlags, BIFF_OBJPIO_LINKED );
                bControl = getFlag( nFlags, BIFF_OBJPIO_CONTROL );
                bCtlsStrm = getFlag( nFlags, BIFF_OBJPIO_CTLSSTREAM );
            }
            break;
            case BIFF_ID_OBJPICTFMLA:
            {
                sal_uInt16 nLinkDataSize = dumpDec< sal_uInt16 >( "link-data-size" );
                sal_uInt32 nLinkDataEnd = rStrm.getRecPos() + nLinkDataSize;
                {
                    IndentGuard aIndGuard2( rOut );
                    sal_uInt16 nFmlaSize = getFormulaDumper().dumpFormulaSize();
                    dumpUnused( 4 );
                    getFormulaDumper().dumpNameFormula( "link", nFmlaSize );
                    if( dumpDec< sal_uInt8 >( "has-class-name", "OBJPICTFMLA-HASCLASSNAME" ) == 3 )
                    {
                        dumpUniString( "class-name", BIFF_STR_SMARTFLAGS );
                        if( ((rStrm.getRecPos() - nSubRecStart) & 1) != 0 )
                            dumpHex< sal_uInt8 >( "padding" );
                    }
                }
                if( rStrm.getRecPos() != nLinkDataEnd )
                    writeEmptyItem( OOX_DUMP_ERRASCII( "link-data-size" ) );
                if( rStrm.getRecPos() < nLinkDataEnd )
                    dumpRemaining( nLinkDataEnd - rStrm.getRecPos() );
                rStrm.seek( nLinkDataEnd );
                if( nLinkDataEnd + 4 <= nSubRecEnd )
                {
                    if( bControl && bCtlsStrm )
                    {
                        sal_uInt32 nPos = dumpHex< sal_uInt32 >( "ctls-stream-pos", "CONV-DEC" );
                        sal_uInt32 nSize = dumpHex< sal_uInt32 >( "ctls-stream-size", "CONV-DEC" );
                        IndentGuard aIndGuard2( rOut );
                        rOut.emptyLine();
                        dumpFormControl( nPos, nSize );
                        rOut.emptyLine();
                    }
                    else
                        dumpHex< sal_uInt32 >( "ole-storage-id" );
                }
                if( bControl && (rStrm.getRecPos() + 8 <= nSubRecEnd) )
                {
                    sal_uInt32 nClassIdSize = dumpDec< sal_uInt32 >( "class-id-size" );
                    if( nClassIdSize > 0 )
                    {
                        IndentGuard aIndGuard2( rOut );
                        sal_uInt32 nClassIdEnd = rStrm.getRecPos() + nClassIdSize;
                        dumpUnicodeArray( "class-id", static_cast< sal_Int32 >( nClassIdSize / 2 ) );
                        rStrm.seek( nClassIdEnd );
                    }
                    sal_uInt16 nCellLinkSize = dumpDec< sal_uInt16 >( "cell-link-size" );
                    if( nCellLinkSize > 0 )
                    {
                        IndentGuard aIndGuard2( rOut );
                        sal_uInt32 nCellLinkEnd = rStrm.getRecPos() + nCellLinkSize;
                        sal_uInt16 nFmlaSize = getFormulaDumper().dumpFormulaSize( "cell-link-fmla-size" );
                        dumpUnused( 4 );
                        getFormulaDumper().dumpNameFormula( "cell-link", nFmlaSize );
                        rStrm.seek( nCellLinkEnd );
                    }
                    sal_uInt16 nSrcRangeSize = dumpDec< sal_uInt16 >( "source-range-size" );
                    if( nSrcRangeSize > 0 )
                    {
                        IndentGuard aIndGuard2( rOut );
                        sal_uInt32 nSrcRangeEnd = rStrm.getRecPos() + nSrcRangeSize;
                        sal_uInt16 nFmlaSize = getFormulaDumper().dumpFormulaSize( "source-range-fmla-size" );
                        dumpUnused( 4 );
                        getFormulaDumper().dumpNameFormula( "source-range", nFmlaSize );
                        rStrm.seek( nSrcRangeEnd );
                    }
                }
            }
            break;
            case BIFF_ID_OBJCMO:
                dumpDec< sal_uInt16 >( "type", "OBJCMO-TYPE" );
                dumpDec< sal_uInt16 >( "id" );
                dumpHex< sal_uInt16 >( "flags", "OBJCMO-FLAGS" );
                dumpUnused( 12 );
            break;
        }
        // remaining undumped data
        if( rStrm.getRecPos() == nSubRecStart )
            dumpRawBinary( nRealRecSize, false );
        else
            dumpRemaining( nSubRecEnd - rStrm.getRecPos() );
        rStrm.seek( nSubRecStart + nRealRecSize );
    }
}

void WorkbookStreamObject::dumpFormControl( sal_uInt32 nStrmPos, sal_uInt32 nStrmSize )
{
    writeHexItem( "stream-pos", nStrmPos, "CONV-DEC" );
    writeHexItem( "stream-size", nStrmSize, "CONV-DEC" );
}

// ============================================================================

PivotCacheStreamObject::PivotCacheStreamObject( const OleStorageObject& rParentStrg, const OUString& rStrmName )
{
    RecordStreamObject::construct( rParentStrg, rStrmName, BIFF8 );
}

PivotCacheStreamObject::~PivotCacheStreamObject()
{
}

void PivotCacheStreamObject::implDumpRecord()
{
    BiffInputStream& rStrm = getBiffStream();
    sal_uInt16 nRecId = rStrm.getRecId();

    switch( nRecId )
    {
        case BIFF_ID_SXDATETIME:
        {
            sal_uInt16 nYear, nMonth;
            sal_uInt8 nDay, nHour, nMin, nSec;
            rStrm >> nYear >> nMonth >> nDay >> nHour >> nMin >> nSec;
            DateTime aDateTime( 0, nSec, nMin, nHour, nDay, nMonth, nYear );
            writeDateTimeItem( "value", aDateTime );
        }
        break;

        case BIFF_ID_SXDB:
            dumpDec< sal_uInt32 >( "source-records" );
            dumpHex< sal_uInt16 >( "stream-id" );
            dumpHex< sal_uInt16 >( "flags", "SXDB-FLAGS" );
            dumpDec< sal_uInt16 >( "block-records" );
            dumpDec< sal_uInt16 >( "standard-field-count" );
            dumpDec< sal_uInt16 >( "total-field-count" );
            dumpUnused( 2 );
            dumpDec< sal_uInt16 >( "database-type", "SXDB-TYPE" );
            dumpUniString( "user-name" );
        break;

        case BIFF_ID_SXFIELD:
            dumpHex< sal_uInt16 >( "flags", "SXFIELD-FLAGS" );
            dumpDec< sal_uInt16 >( "group-child-field" );
            dumpDec< sal_uInt16 >( "group-base-field" );
            dumpDec< sal_uInt16 >( "visible-items" );
            dumpDec< sal_uInt16 >( "group-items" );
            dumpDec< sal_uInt16 >( "base-items" );
            dumpDec< sal_uInt16 >( "original-items" );
            if( rStrm.getRecLeft() >= 3 )
                dumpUniString( "item-name" );
        break;

        case BIFF_ID_SXSTRING:
            dumpUniString( "value" );
        break;
    }
}

// ============================================================================

VbaProjectStreamObject::VbaProjectStreamObject( const OleStorageObject& rParentStrg )
{
    OleStreamObject::construct( rParentStrg, CREATE_OUSTRING( "PROJECT" ) );
}

void VbaProjectStreamObject::implDumpBody()
{
    dumpTextStream( osl_getThreadTextEncoding() );
}

// ============================================================================
// ============================================================================

PivotCacheStorageObject::PivotCacheStorageObject( const OleStorageObject& rParentStrg )
{
    OleStorageObject::construct( rParentStrg, CREATE_OUSTRING( "_SX_DB_CUR" ) );
}

void PivotCacheStorageObject::implDumpBody()
{
    for( OleStorageIterator aIt( *this ); aIt.isValid(); ++aIt )
        if( aIt.isStream() )
            PivotCacheStreamObject( *this, aIt.getName() ).dump();
}

// ============================================================================

VbaProjectStorageObject::VbaProjectStorageObject( const OleStorageObject& rParentStrg )
{
    OleStorageObject::construct( rParentStrg, CREATE_OUSTRING( "_VBA_PROJECT_CUR" ) );
}

void VbaProjectStorageObject::implDumpBody()
{
    VbaProjectStreamObject( *this ).dump();
    VbaStorageObject( *this ).dump();
}

// ============================================================================

VbaStorageObject::VbaStorageObject( const OleStorageObject& rParentStrg )
{
    OleStorageObject::construct( rParentStrg, CREATE_OUSTRING( "VBA" ) );
}

void VbaStorageObject::implDumpBody()
{
    OleStreamObject( *this, CREATE_OUSTRING( "dir" ) ).dump();
    OleStreamObject( *this, CREATE_OUSTRING( "ThisWorkbook" ) ).dump();
    OleStreamObject( *this, CREATE_OUSTRING( "_VBA_PROJECT" ) ).dump();
}

// ============================================================================

RootStorageObject::RootStorageObject( const ObjectBase& rParent )
{
    OleStorageObject::construct( rParent );
}

void RootStorageObject::implDumpBody()
{
    extractStorageToFileSystem();
    WorkbookStreamObject( *this, CREATE_OUSTRING( "Workbook" ) ).dump();
    WorkbookStreamObject( *this, CREATE_OUSTRING( "Book" ) ).dump();
    OlePropertyStreamObject( *this, CREATE_OUSTRING( "\005SummaryInformation" ) ).dump();
    OlePropertyStreamObject( *this, CREATE_OUSTRING( "\005DocumentSummaryInformation" ) ).dump();
    OleStreamObject( *this, CREATE_OUSTRING( "Ctls" ) ).dump();
    PivotCacheStorageObject( *this ).dump();
    VbaProjectStorageObject( *this ).dump();
}

// ============================================================================
// ============================================================================

Dumper::Dumper( const FilterBase& rFilter )
{
    ConfigRef xCfg( new Config( "OOO_BIFFDUMPER" ) );
    DumperBase::construct( rFilter, xCfg );
}

void Dumper::implDumpBody()
{
    WorkbookStreamObject( *this ).dump();
    RootStorageObject( *this ).dump();
}

// ============================================================================

} // namespace biff
} // namespace dump
} // namespace oox

#endif

