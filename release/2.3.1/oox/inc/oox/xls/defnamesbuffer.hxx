/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: defnamesbuffer.hxx,v $
 *
 *  $Revision: 1.1.2.8 $
 *
 *  last change: $Author: dr $ $Date: 2007/08/21 08:09:05 $
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

#ifndef OOX_XLS_DEFINEDNAMESBUFFER_HXX
#define OOX_XLS_DEFINEDNAMESBUFFER_HXX

#include "oox/core/containerhelper.hxx"
#include "oox/xls/globaldatahelper.hxx"

namespace com { namespace sun { namespace star {
    namespace sheet { class XNamedRange; }
} } }

namespace oox { namespace core {
    class AttributeList;
} }

namespace oox {
namespace xls {

// ============================================================================

// flags
const sal_uInt16 BIFF_NAME_DEFAULT          = 0x0000;
const sal_uInt16 BIFF_NAME_HIDDEN           = 0x0001;
const sal_uInt16 BIFF_NAME_FUNC             = 0x0002;
const sal_uInt16 BIFF_NAME_VBNAME           = 0x0004;
const sal_uInt16 BIFF_NAME_MACRO            = 0x0008;
const sal_uInt16 BIFF_NAME_CALCEXP          = 0x0010;
const sal_uInt16 BIFF_NAME_BUILTIN          = 0x0020;
const sal_uInt16 BIFF_NAME_FGROUPMASK       = 0x0FC0;
const sal_uInt16 BIFF_NAME_BIG              = 0x1000;

const sal_uInt8 BIFF_NAME2_FUNC             = 0x02;     /// BIFF2 function/command flag.

const sal_uInt16 BIFF_NAME_GLOBAL           = 0;        /// 0 = Globally defined name.

// codes for built-in names
const sal_Unicode BIFF_BUILTIN_CONSOLIDATEAREA  = '\x00';
const sal_Unicode BIFF_BUILTIN_AUTOOPEN         = '\x01';
const sal_Unicode BIFF_BUILTIN_AUTOCLOSE        = '\x02';
const sal_Unicode BIFF_BUILTIN_EXTRACT          = '\x03';
const sal_Unicode BIFF_BUILTIN_DATABASE         = '\x04';
const sal_Unicode BIFF_BUILTIN_CRITERIA         = '\x05';
const sal_Unicode BIFF_BUILTIN_PRINTAREA        = '\x06';
const sal_Unicode BIFF_BUILTIN_PRINTTITLES      = '\x07';
const sal_Unicode BIFF_BUILTIN_RECORDER         = '\x08';
const sal_Unicode BIFF_BUILTIN_DATAFORM         = '\x09';
const sal_Unicode BIFF_BUILTIN_AUTOACTIVATE     = '\x0A';
const sal_Unicode BIFF_BUILTIN_AUTODEACTIVATE   = '\x0B';
const sal_Unicode BIFF_BUILTIN_SHEETTITLE       = '\x0C';
const sal_Unicode BIFF_BUILTIN_FILTERDATABASE   = '\x0D';
const sal_Unicode BIFF_BUILTIN_UNKNOWN          = '\x0E';

// ============================================================================

struct OoxDefinedNameData
{
    ::rtl::OUString     maName;         /// The original name.
    ::rtl::OUString     maFormula;      /// The formula string.
    sal_Int32           mnSheet;        /// Sheet index for local names.
    bool                mbMacro;        /// True = Macro name (VBasic or sheet macro).
    bool                mbFunction;     /// True = function, false = command.
    bool                mbVBName;       /// True = VBasic macro, false = sheet macro.
    bool                mbHidden;       /// True = name hidden in UI.

    explicit            OoxDefinedNameData();
};

// ============================================================================

class DefinedName : public GlobalDataHelper
{
public:
    explicit            DefinedName( const GlobalDataHelper& rGlobalData, sal_Int32 nLocalSheet );

    /** Sets the attributes for this defined name from the passed attribute set. */
    void                importDefinedName( const ::oox::core::AttributeList& rAttribs );
    /** Sets the formula string from the body of the definedName element. */
    void                setFormula( const ::rtl::OUString& rFormula );
    /** Imports the defined name from a NAME record in the passed BIFF stream. */
    void                importName( BiffInputStream& rStrm );

    /** Creates a defined name in the Calc document. */
    void                createNameObject();
    /** Converts the formula string or BIFF token array for this defined name. */
    void                convertFormula();

    /** Returns true, if this defined name is global in the document. */
    inline bool         isGlobalName() const { return maOoxData.mnSheet < 0; }
    /** Returns true, if this defined name is a special builtin name. */
    inline bool         isBuiltinName() const { return mcBuiltinId != BIFF_BUILTIN_UNKNOWN; }
    /** Returns true, if this defined name is a macro function call. */
    inline bool         isMacroFunc( bool bVBName ) const { return maOoxData.mbMacro && maOoxData.mbFunction && (maOoxData.mbVBName == bVBName); }

    /** Returns the original name as imported from or exported to the file. */
    inline const ::rtl::OUString& getOoxName() const { return maOoxData.maName; }
    /** Returns the token index used in API token arrays (com.sun.star.sheet.FormulaToken). */
    inline sal_Int32    getTokenIndex() const { return mnTokenIndex; }
    /** Returns the 0-based sheet index for local names, or -1 for global names. */
    inline sal_Int32    getSheetIndex() const { return maOoxData.mnSheet; }

private:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XNamedRange > XNamedRangeRef;

    OoxDefinedNameData  maOoxData;          /// OOX data for this defined name.
    ::rtl::OUString     maFinalName;        /// Final name used in the Calc document.
    XNamedRangeRef      mxNamedRange;       /// XNamedRange interface of the defined name.
    sal_Int32           mnTokenIndex;       /// Name index used in API token array.
    sal_Unicode         mcBuiltinId;        /// Identifier for built-in defined names.
    BiffInputStream*    mpStrm;             /// Cached BIFF stream for formula import.
    sal_Int64           mnRecHandle;        /// Cached BIFF record handle for formula import.
    sal_uInt32          mnRecPos;           /// Cached BIFF record position for formula import.
    sal_uInt16          mnFmlaSize;         /// Cached BIFF formula size for formula import.
};

typedef ::boost::shared_ptr< DefinedName > DefinedNameRef;

// ============================================================================

class DefinedNamesBuffer : public GlobalDataHelper
{
public:
    explicit            DefinedNamesBuffer( const GlobalDataHelper& rGlobalData );

    /** Creates and returns a defined name on-the-fly in the Calc document.
        The name will not be buffered in this defined names buffer.
        @param orName  (in/out-parameter) Returns the resulting used name. */
    ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XNamedRange >
                        createDefinedName( ::rtl::OUString& orName, sal_Int32 nNameFlags = 0 ) const;

    /** Returns the index of the passed defined name used in formula token arrays. */
    sal_Int32           getTokenIndex( const ::com::sun::star::uno::Reference< ::com::sun::star::sheet::XNamedRange >& rxNamedRange ) const;

    /** Sets the current sheet index for files with local defined names, e.g.
        BIFF4 workspaces. All created names initially will contain this index. */
    void                setLocalSheetIndex( sal_Int32 nLocalSheet );
    /** Imports a defined name from the passed attribute set. */
    DefinedNameRef      importDefinedName( const ::oox::core::AttributeList& rAttribs );
    /** Imports a defined name from a NAME record in the passed BIFF stream. */
    void                importName( BiffInputStream& rStrm );

    /** Creates all defined names in the document. */
    void                finalizeImport();

    /** Returns a defined name by one-based BIFF index. */
    DefinedNameRef      getByBiffIndex( sal_uInt16 nIndex ) const;
    /** Returns a defined name by its OOX name.
        @param nSheet  The sheet index for local names or -1 for global names.
            If no local name is found, tries to find a matching global name.
        @return  Reference to the defined name or empty reference. */
    DefinedNameRef      getByOoxName( const ::rtl::OUString& rOoxName, sal_Int32 nSheet = -1 ) const;

private:
    DefinedNameRef      createDefinedName();

private:
    typedef ::oox::core::RefVector< DefinedName > DefNameVec;

    const ::rtl::OUString maTokenIndexProp;
    DefNameVec          maDefNames;
    sal_Int32           mnLocalSheet;
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

