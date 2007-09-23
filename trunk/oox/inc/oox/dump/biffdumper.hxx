/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffdumper.hxx,v $
 *
 *  $Revision: 1.1.2.11 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/13 13:43:01 $
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

#ifndef OOX_DUMP_BIFFDUMPER_HXX
#define OOX_DUMP_BIFFDUMPER_HXX

#include <stack>
#include "oox/dump/dumperbase.hxx"
#include "oox/dump/dffdumper.hxx"
#include "oox/dump/olestoragedumper.hxx"
#include "oox/xls/biffhelper.hxx"
#include "oox/xls/sharedstringsbuffer.hxx"

#if OOX_INCLUDE_DUMPER

namespace oox { namespace xls {
    class BiffInputStream;
    class FunctionProvider;
} }

namespace oox {
namespace dump {
namespace biff {

// ============================================================================
// ============================================================================

struct Address
{
    sal_Int32           mnCol;
    sal_Int32           mnRow;
    inline explicit     Address() : mnCol( 0 ), mnRow( 0 ) {}
    inline explicit     Address( sal_Int32 nCol, sal_Int32 nRow ) : mnCol( nCol ), mnRow( nRow ) {}
    void                read( ::oox::xls::BiffInputStream& rStrm, bool bCol16Bit = true, bool bRow32Bit = false );
};

// ----------------------------------------------------------------------------

struct Range
{
    Address             maFirst;
    Address             maLast;
    inline explicit     Range() {}
    void                read( ::oox::xls::BiffInputStream& rStrm, bool bCol16Bit = true, bool bRow32Bit = false );
};

// ----------------------------------------------------------------------------

struct RangeList : public ::std::vector< Range >
{
    inline explicit     RangeList() {}
    void                read( ::oox::xls::BiffInputStream& rStrm, bool bCol16Bit = true, bool bRow32Bit = false );
};

// ============================================================================

struct TokenAddress : public Address
{
    bool                mbRelCol;
    bool                mbRelRow;
    inline explicit     TokenAddress() : mbRelCol( false ), mbRelRow( false ) {}
};

// ----------------------------------------------------------------------------

struct TokenRange
{
    TokenAddress        maFirst;
    TokenAddress        maLast;
    inline explicit     TokenRange() {}
};

// ============================================================================

class StringHelper : public ::oox::dump::StringHelper
{
public:
    static void         appendAddrCol( ::rtl::OUStringBuffer& rStr, sal_Int32 nCol, bool bRel );
    static void         appendAddrRow( ::rtl::OUStringBuffer& rStr, sal_Int32 nRow, bool bRel );
    static void         appendAddrName( ::rtl::OUStringBuffer& rStr, sal_Unicode cPrefix, sal_Int32 nColRow, bool bRel );

    static void         appendAddress( ::rtl::OUStringBuffer& rStr, const Address& rPos );
    static void         appendRange( ::rtl::OUStringBuffer& rStr, const Range& rRange );
    static void         appendRangeList( ::rtl::OUStringBuffer& rStr, const RangeList& rRanges );

    static void         appendAddress( ::rtl::OUStringBuffer& rStr, const TokenAddress& rPos, bool bNameMode );
    static void         appendRange( ::rtl::OUStringBuffer& rStr, const TokenRange& rRange, bool bNameMode );
};

// ============================================================================
// ============================================================================

class RecordHeaderObject : public RecordHeaderBase
{
public:
    explicit            RecordHeaderObject( const InputObjectBase& rParent );

    inline bool         isMergeContRec() const { return mbMergeContRec; }
    inline bool         hasRecName( sal_uInt16 nRecId ) const { return getRecNames()->hasName( nRecId ); }

    void                dumpRecordHeader( ::oox::xls::BiffInputStream& rStrm );

protected:
    virtual bool        implIsValid() const;

private:
    bool                mbMergeContRec;
};

typedef ::boost::shared_ptr< RecordHeaderObject > RecHeaderObjectRef;

// ============================================================================
// ============================================================================

class BiffConfig : public Config
{
public:
    explicit            BiffConfig( const Config& rParent, ::oox::xls::BiffType eBiff );

protected:
    virtual bool        implIsValid() const;
    virtual NameListRef implGetNameList( const ::rtl::OUString& rKey ) const;

private:
    ::oox::xls::BiffType meBiff;
};

typedef ::boost::shared_ptr< BiffConfig > BiffConfigRef;

// ============================================================================

class BiffSharedData : public Base
{
public:
    explicit            BiffSharedData( ::oox::xls::BiffType eBiff );
    virtual             ~BiffSharedData();

    inline ::oox::xls::BiffType getBiff() const { return meBiff; }

    inline rtl_TextEncoding getTextEncoding() const { return meTextEnc; }
    void                setTextEncoding( rtl_TextEncoding eTextEnc );

    sal_uInt16          getFontCount() const;
    rtl_TextEncoding    getFontEncoding( sal_uInt16 nFontId ) const;
    void                appendFontEncoding( rtl_TextEncoding eFontEnc );

    sal_uInt16          getXfCount() const;
    rtl_TextEncoding    getXfEncoding( sal_uInt16 nXfId ) const;
    void                appendXfFontId( sal_uInt16 nFontId );

protected:
    virtual bool        implIsValid() const;

private:
    typedef ::std::vector< rtl_TextEncoding >   TextEncVec;
    typedef ::std::vector< sal_uInt16 >         FontIdVec;

    TextEncVec          maFontEncs;
    FontIdVec           maXfFontIds;
    ::oox::xls::BiffType meBiff;
    rtl_TextEncoding    meTextEnc;
};

typedef ::boost::shared_ptr< BiffSharedData > BiffSharedDataRef;

// ============================================================================

class BiffObjectBase : public WrappedStreamObject
{
public:
    virtual             ~BiffObjectBase();

    inline BiffSharedData& getBiffData() const { return *mxBiffData; }
    inline ::oox::xls::BiffInputStream& getBiffStream() const { return *mxStrm; }
    inline ::oox::xls::BiffType getBiff() const { return mxBiffData->getBiff(); }

protected:
    explicit            BiffObjectBase();
    void                construct( const ObjectBase& rParent, BinaryInputStreamRef xStrm, ::oox::xls::BiffType eBiff );
    void                construct( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName, ::oox::xls::BiffType eBiff );
    void                construct( const BiffObjectBase& rParent );

    virtual bool        implIsValid() const;
    virtual void        implDumpExtendedHeader();
    virtual ConfigRef   implReconstructConfig();
    virtual InputRef    implReconstructInput();

    ::rtl::OUString     getErrorName( sal_uInt8 nErrCode ) const;

    // ------------------------------------------------------------------------

    double              writeRkItem( const sal_Char* pcName, sal_Int32 nRk );
    void                writeBooleanItem( const sal_Char* pcName, sal_uInt8 nBool );
    void                writeErrorCodeItem( const sal_Char* pcName, sal_uInt8 nErrCode );

    void                writeColIndexItem( const sal_Char* pcName, sal_Int32 nCol );
    void                writeColRangeItem( const sal_Char* pcName, sal_Int32 nCol1, sal_Int32 nCol2 );
    void                writeRowIndexItem( const sal_Char* pcName, sal_Int32 nRow );
    void                writeRowRangeItem( const sal_Char* pcName, sal_Int32 nRow1, sal_Int32 nRow2 );

    void                writeAddressItem( const sal_Char* pcName, const Address& rPos );
    void                writeRangeItem( const sal_Char* pcName, const Range& rRange );
    void                writeRangeListItem( const sal_Char* pcName, const RangeList& rRanges );

    void                writeRichStringFontIds( const ::oox::xls::BiffRichStringFontIdVec& rFontIds );

    template< typename Type >
    void                writeRectItem( const sal_Char* pcName,
                            Type nLeft, Type nTop, Type nWidth, Type nHeight,
                            const NameListWrapper& rListWrp = NO_LIST,
                            FormatType eFmtType = FORMATTYPE_DEC );

    // ------------------------------------------------------------------------

    ::rtl::OUString     dumpByteString(
                            const sal_Char* pcName,
                            ::oox::xls::BiffStringFlags nFlags = ::oox::xls::BIFF_STR_DEFAULT,
                            rtl_TextEncoding eDefaultTextEnc = RTL_TEXTENCODING_DONTKNOW );
    ::rtl::OUString     dumpUniString(
                            const sal_Char* pcName,
                            ::oox::xls::BiffStringFlags nFlags = ::oox::xls::BIFF_STR_DEFAULT );
    ::rtl::OUString     dumpString(
                            const sal_Char* pcName,
                            ::oox::xls::BiffStringFlags nByteFlags = ::oox::xls::BIFF_STR_DEFAULT,
                            ::oox::xls::BiffStringFlags nUniFlags = ::oox::xls::BIFF_STR_DEFAULT,
                            rtl_TextEncoding eDefaultTextEnc = RTL_TEXTENCODING_DONTKNOW );
    ::rtl::OUString     dumpOleString(
                            const sal_Char* pcName,
                            sal_Int32 nCharCount,
                            bool bUnicode );
    ::rtl::OUString     dumpOleString(
                            const sal_Char* pcName,
                            bool bUnicode );
    ::rtl::OUString     dumpNullString(
                            const sal_Char* pcName,
                            bool bUnicode );

    double              dumpRk( const sal_Char* pcName = 0 );
    sal_uInt8           dumpBoolean( const sal_Char* pcName = 0 );
    sal_uInt8           dumpErrorCode( const sal_Char* pcName = 0 );

    sal_Int32           dumpRgbColor( const sal_Char* pcName = 0 );
    rtl_TextEncoding    dumpCodePage( const sal_Char* pcName = 0 );
    void                dumpFormulaResult( const sal_Char* pcName = 0 );

    sal_Int32           dumpColIndex( const sal_Char* pcName = 0, bool bCol16Bit = true );
    sal_Int32           dumpColRange( const sal_Char* pcName = 0, bool bCol16Bit = true );
    sal_Int32           dumpRowIndex( const sal_Char* pcName = 0, bool bRow32Bit = false );
    sal_Int32           dumpRowRange( const sal_Char* pcName = 0, bool bRow32Bit = false );

    Address             dumpAddress( const sal_Char* pcName = 0, bool bCol16Bit = true, bool bRow32Bit = false );
    Range               dumpRange( const sal_Char* pcName = 0, bool bCol16Bit = true, bool bRow32Bit = false );
    void                dumpRangeList( const sal_Char* pcName = 0, bool bCol16Bit = true, bool bRow32Bit = false );

    void                dumpConstArrayHeader( sal_uInt32& rnCols, sal_uInt32& rnRows );
    ::rtl::OUString     dumpConstValue();

    template< typename Type >
    void                dumpRect( const sal_Char* pcName,
                            const NameListWrapper& rListWrp = NO_LIST,
                            FormatType eFmtType = FORMATTYPE_DEC );

    // ------------------------------------------------------------------------

    using               WrappedStreamObject::construct;

private:
    void                constructBiffObjBase( ::oox::xls::BiffType eBiff );

private:
    typedef ::boost::shared_ptr< ::oox::xls::BiffInputStream > BiffInputStreamRef;

    BiffConfigRef       mxBiffCfg;
    BiffSharedDataRef   mxBiffData;
    BiffInputStreamRef  mxStrm;
    NameListRef         mxBoolean;
    NameListRef         mxErrCodes;
    NameListRef         mxConstType;
    NameListRef         mxResultType;
};

// ----------------------------------------------------------------------------

template< typename Type >
void BiffObjectBase::writeRectItem( const sal_Char* pcName,
        Type nLeft, Type nTop, Type nWidth, Type nHeight,
        const NameListWrapper& rListWrp, FormatType eFmtType )
{
    MultiItemsGuard aMultiGuard( out() );
    writeEmptyItem( pcName );
    writeValueItem( "x-pos", nLeft, eFmtType, rListWrp );
    writeValueItem( "y-pos", nTop, eFmtType, rListWrp );
    writeValueItem( "x-size", nWidth, eFmtType, rListWrp );
    writeValueItem( "y-size", nHeight, eFmtType, rListWrp );
}

template< typename Type >
void BiffObjectBase::dumpRect( const sal_Char* pcName,
        const NameListWrapper& rListWrp, FormatType eFmtType )
{
    Type nLeft, nTop, nWidth, nHeight;
    *mxStrm >> nLeft >> nTop >> nWidth >> nHeight;
    writeRectItem( pcName, nLeft, nTop, nWidth, nHeight, rListWrp, eFmtType );
}

// ============================================================================
// ============================================================================

class SimpleObjectBase : public BiffObjectBase
{
public:
    virtual             ~SimpleObjectBase();

protected:
    inline explicit     SimpleObjectBase() {}
    void                construct( const BiffObjectBase& rParent );

    virtual void        implDumpHeader();
    virtual void        implDumpFooter();

    using               BiffObjectBase::construct;
};

// ============================================================================
// ============================================================================

/** Stack to create a human readable formula string from a UPN token array. */
class FormulaStack
{
public:
    explicit            FormulaStack();

    inline const ::rtl::OUString& getFormulaString() const { return getString( maFmlaStack ); }
    inline const ::rtl::OUString& getClassesString() const { return getString( maClassStack ); }

    void                pushOperand( const StringWrapper& rOp, const ::rtl::OUString& rTokClass );
    void                pushOperand( const StringWrapper& rOp );
    void                pushUnaryOp( const StringWrapper& rLOp, const StringWrapper& rROp );
    void                pushBinaryOp( const StringWrapper& rOp );
    void                pushFuncOp( const StringWrapper& rFunc, const ::rtl::OUString& rTokClass, sal_uInt8 nParamCount );

    inline void         setError() { mbError = true; }
    void                replaceOnTop( const ::rtl::OUString& rOld, const ::rtl::OUString& rNew );

private:
    typedef ::std::stack< ::rtl::OUString > StringStack;

    inline bool         check( bool bCond ) { return (mbError |= !bCond) == false; }

    const ::rtl::OUString& getString( const StringStack& rStack ) const;
    void                pushUnaryOp( StringStack& rStack, const ::rtl::OUString& rLOp, const ::rtl::OUString& rROp );
    void                pushBinaryOp( StringStack& rStack, const ::rtl::OUString& rOp );
    void                pushFuncOp( StringStack& rStack, const ::rtl::OUString& rOp, sal_uInt8 nParamCount );

private:
    StringStack         maFmlaStack;
    StringStack         maClassStack;
    bool                mbError;
};

// ============================================================================

class FormulaObject : public SimpleObjectBase
{
public:
    explicit            FormulaObject( const BiffObjectBase& rParent );
    virtual             ~FormulaObject();

    sal_uInt16          readFormulaSize();
    sal_uInt16          dumpFormulaSize( const sal_Char* pcName = 0 );

    void                dumpCellFormula( const sal_Char* pcName, sal_uInt16 nSize );
    void                dumpCellFormula( const sal_Char* pcName = 0 );
    void                dumpNameFormula( const sal_Char* pcName, sal_uInt16 nSize );
    void                dumpNameFormula( const sal_Char* pcName = 0 );

protected:
    virtual void        implDumpHeader();
    virtual void        implDumpBody();

private:
    void                constructFmlaObj();

    void                dumpFormula( const sal_Char* pcName, sal_uInt16 nSize, bool bNameMode );
    void                dumpFormula( const sal_Char* pcName, bool bNameMode );

    TokenAddress        createTokenAddress( sal_uInt16 nCol, sal_uInt16 nRow, bool bRelC, bool bRelR, bool bNameMode ) const;
    ::rtl::OUString     createFunc( sal_uInt16 nFuncIdx ) const;
    ::rtl::OUString     createRef( const ::rtl::OUString& rData ) const;
    ::rtl::OUString     createName( sal_uInt16 nNameIdx ) const;
    ::rtl::OUString     createPlaceHolder( size_t nIdx ) const;
    ::rtl::OUString     createPlaceHolder() const;

    void                writeTokenAddressItem( const sal_Char* pcName, const TokenAddress& rPos, bool bNameMode );
    void                writeTokenAddress3dItem( const sal_Char* pcName, const ::rtl::OUString& rRef, const TokenAddress& rPos, bool bNameMode );
    void                writeTokenRangeItem( const sal_Char* pcName, const TokenRange& rRange, bool bNameMode );
    void                writeTokenRange3dItem( const sal_Char* pcName, const ::rtl::OUString& rRef, const TokenRange& rRange, bool bNameMode );

    sal_uInt16          dumpTokenFuncIdx();
    sal_uInt16          dumpTokenCol( const sal_Char* pcName, bool& rbRelC, bool& rbRelR );
    sal_uInt16          dumpTokenRow( const sal_Char* pcName, bool& rbRelC, bool& rbRelR );
    TokenAddress        dumpTokenAddress( bool bNameMode );
    TokenRange          dumpTokenRange( bool bNameMode );
    ::rtl::OUString     dumpTokenRefIdx();
    void                dumpTokenRefTabIdxs();

    void                dumpIntToken();
    void                dumpDoubleToken();
    void                dumpStringToken();
    void                dumpBoolToken();
    void                dumpErrorToken();
    void                dumpMissArgToken();

    void                dumpArrayToken( const ::rtl::OUString& rTokClass );
    void                dumpNameToken( const ::rtl::OUString& rTokClass );
    void                dumpNameXToken( const ::rtl::OUString& rTokClass );
    void                dumpRefToken( const ::rtl::OUString& rTokClass, bool bNameMode );
    void                dumpAreaToken( const ::rtl::OUString& rTokClass, bool bNameMode );
    void                dumpRefErrToken( const ::rtl::OUString& rTokClass, bool bArea );
    void                dumpRef3dToken( const ::rtl::OUString& rTokClass, bool bNameMode );
    void                dumpArea3dToken( const ::rtl::OUString& rTokClass, bool bNameMode );
    void                dumpRefErr3dToken( const ::rtl::OUString& rTokClass, bool bArea );
    void                dumpMemFuncToken( const ::rtl::OUString& rTokClass );
    void                dumpMemAreaToken( const ::rtl::OUString& rTokClass, bool bAddData );

    void                dumpExpToken( const StringWrapper& rName );
    void                dumpUnaryOpToken( const StringWrapper& rLOp, const StringWrapper& rROp );
    void                dumpBinaryOpToken( const StringWrapper& rOp );
    void                dumpFuncToken( const ::rtl::OUString& rTokClass );
    void                dumpFuncVarToken( const ::rtl::OUString& rTokClass );
    void                dumpCmdToken( const ::rtl::OUString& rTokClass );

    void                dumpSheetToken();
    void                dumpEndSheetToken();
    bool                dumpAttrToken();

    bool                dumpNlrToken();
    void                dumpNlrErrToken();
    void                dumpNlrColRowToken( const ::rtl::OUString& rTokClass, bool bAddData );
    void                dumpNlrRangeToken( const ::rtl::OUString& rTokClass, bool bAddData );
    void                dumpNlrRangeErrToken();

    void                dumpAddTokenData();
    void                dumpAddDataNlr( size_t nIdx );
    void                dumpAddDataArray( size_t nIdx );
    void                dumpAddDataMemArea( size_t nIdx );

private:
    enum AddDataType { ADDDATA_NLR, ADDDATA_ARRAY, ADDDATA_MEMAREA };

    typedef ::boost::shared_ptr< FormulaStack >                 FormulaStackRef;
    typedef ::boost::shared_ptr< ::oox::xls::FunctionProvider > FuncProvRef;
    typedef ::std::vector< AddDataType >                        AddDataTypeVec;

    NameListRef         mxTokens;
    NameListRef         mxClasses;
    NameListRef         mxFuncNames;
    NameListRef         mxParamCnt;
    NameListRef         mxRelFlags;
    NameListRef         mxNlrTypes;
    NameListRef         mxAttrTypes;
    NameListRef         mxSpTypes;
    sal_Int32           mnColCount;
    sal_Int32           mnRowCount;

    FormulaStackRef     mxStack;
    FuncProvRef         mxFuncProv;
    AddDataTypeVec      maAddData;
    ::rtl::OUString     maRefPrefix;
    const sal_Char*     mpcName;
    sal_uInt16          mnSize;
    bool                mbNameMode;
};

typedef ::boost::shared_ptr< FormulaObject > FormulaObjectRef;

// ============================================================================
// ============================================================================

class RecordStreamObject : public BiffObjectBase
{
public:
    explicit            RecordStreamObject( const ObjectBase& rParent, BinaryInputStreamRef xStrm, ::oox::xls::BiffType eBiff = ::oox::xls::BIFF_UNKNOWN );
    explicit            RecordStreamObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName, ::oox::xls::BiffType eBiff = ::oox::xls::BIFF_UNKNOWN );
    virtual             ~RecordStreamObject();

protected:
    inline explicit     RecordStreamObject() {}
    void                construct( const ObjectBase& rParent, BinaryInputStreamRef xStrm, ::oox::xls::BiffType eBiff );
    void                construct( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName, ::oox::xls::BiffType eBiff );

    virtual bool        implIsValid() const;
    virtual void        implDumpBody();

    virtual void        implPreProcessRecord();
    virtual void        implDumpRecord();
    virtual void        implPostProcessRecord();

    inline RecordHeaderObject& getRecordHeader() const { return *mxHdrObj; }
    inline FormulaObject& getFormulaDumper() const { return *mxFmlaObj; }
    inline DffDumpObject& getDffDumper() const { return *mxDffObj; }

    sal_uInt16          dumpRepeatedRecId();

    using               BiffObjectBase::construct;

private:
    void                constructRecStrmObj();

    void                dumpRecordBody();
    void                dumpSimpleRecord( const ::rtl::OUString& rRecData );

private:
    RecHeaderObjectRef  mxHdrObj;
    FormulaObjectRef    mxFmlaObj;
    DffDumpObjectRef    mxDffObj;
    NameListRef         mxSimpleRecs;
};

// ============================================================================

class WorkbookStreamObject : public RecordStreamObject
{
public:
    explicit            WorkbookStreamObject( const DumperBase& rDumper );
    explicit            WorkbookStreamObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );
    virtual             ~WorkbookStreamObject();

protected:
    virtual void        implPreProcessRecord();
    virtual void        implDumpRecord();
    virtual void        implPostProcessRecord();

private:
    void                constructWbStrmObj();

    ::rtl::OUString     createFontName( const ::rtl::OUString& rName, sal_uInt16 nHeight, bool bBold, bool bItalic ) const;

    sal_uInt16          dumpPatternIdx( const sal_Char* pcName = 0, bool b16Bit = true );
    sal_uInt16          dumpColorIdx( const sal_Char* pcName = 0, bool b16Bit = true );
    sal_uInt16          dumpFontIdx( const sal_Char* pcName = 0, bool b16Bit = true );
    sal_uInt16          dumpFormatIdx( const sal_Char* pcName = 0 );
    sal_uInt16          dumpXfIdx( const sal_Char* pcName = 0, bool bBiff2Style = false );

    sal_uInt16          dumpCellHeader( bool bBiff2Style = false );
    void                dumpBoolErr();

    void                dumpFontRec();
    void                dumpFormatRec();
    void                dumpXfRec();

    void                dumpObjRec();
    void                dumpObjRecBiff5();
    void                dumpObjRecBiff8();
    void                dumpFormControl( sal_uInt32 nStrmPos, sal_uInt32 nStrmSize );

private:
    NameListRef         mxColors;
    NameListRef         mxBorderStyles;
    NameListRef         mxFillPatterns;
    NameListRef         mxFontNames;
    NameListRef         mxFormats;
    sal_uInt16          mnFormatIdx;
    sal_uInt16          mnPTRowFields;
    sal_uInt16          mnPTColFields;
    sal_uInt16          mnPTSxliIdx;
    bool                mbHasCodePage;
};

// ============================================================================

class PivotCacheStreamObject : public RecordStreamObject
{
public:
    explicit            PivotCacheStreamObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );
    virtual             ~PivotCacheStreamObject();

protected:
    virtual void        implDumpRecord();
};

// ============================================================================

class VbaProjectStreamObject : public OleStreamObject
{
public:
    explicit            VbaProjectStreamObject( const OleStorageObject& rParentStrg );

protected:
    virtual void        implDumpBody();
};

// ============================================================================
// ============================================================================

class PivotCacheStorageObject : public OleStorageObject
{
public:
    explicit            PivotCacheStorageObject( const OleStorageObject& rParentStrg );

protected:
    virtual void        implDumpBody();
};

// ============================================================================

class VbaProjectStorageObject : public OleStorageObject
{
public:
    explicit            VbaProjectStorageObject( const OleStorageObject& rParentStrg );

protected:
    virtual void        implDumpBody();
};

// ============================================================================

class VbaStorageObject : public OleStorageObject
{
public:
    explicit            VbaStorageObject( const OleStorageObject& rParentStrg );

protected:
    virtual void        implDumpBody();
};

// ============================================================================

class RootStorageObject : public OleStorageObject
{
public:
    explicit            RootStorageObject( const ObjectBase& rParent );

protected:
    virtual void        implDumpBody();
};

// ============================================================================
// ============================================================================

class Dumper : public DumperBase
{
public:
    explicit            Dumper( const ::oox::core::FilterBase& rFilter );

protected:
    virtual void        implDumpBody();
};

// ============================================================================
// ============================================================================

} // namespace biff
} // namespace dump
} // namespace oox

#endif
#endif

