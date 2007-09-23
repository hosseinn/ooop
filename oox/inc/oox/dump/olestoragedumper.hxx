/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: olestoragedumper.hxx,v $
 *
 *  $Revision: 1.1.2.5 $
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

#ifndef OOX_DUMP_OLESTORAGEDUMPER_HXX
#define OOX_DUMP_OLESTORAGEDUMPER_HXX

#include "oox/core/storagebase.hxx"
#include "oox/dump/dumperbase.hxx"

#if OOX_INCLUDE_DUMPER

namespace com { namespace sun { namespace star {
    namespace io { class XInputStream; }
} } }

namespace oox {
namespace dump {

// ============================================================================
// ============================================================================

class OleStorageObject : public ObjectBase
{
public:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::io::XInputStream > XInputStreamRef;

public:
    explicit            OleStorageObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrgName );
    explicit            OleStorageObject( const ObjectBase& rParent, ::oox::core::StorageRef xStrg );
    explicit            OleStorageObject( const ObjectBase& rParent, const XInputStreamRef& rxRootStrm );
    explicit            OleStorageObject( const ObjectBase& rParent );
    virtual             ~OleStorageObject();

    inline ::oox::core::StorageRef getStorage() const { return mxStrg; }
    inline const ::rtl::OUString& getStoragePath() const { return maPath; }
    inline const ::rtl::OUString& getStorageName() const { return mxStrg->getName(); }
    ::rtl::OUString     getFullName() const;

    void                extractStorageToFileSystem();

protected:
    inline explicit     OleStorageObject() {}
    void                construct( const ObjectBase& rParent, ::oox::core::StorageRef xStrg, const ::rtl::OUString& rPath );
    void                construct( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrgName );
    void                construct( const ObjectBase& rParent, const XInputStreamRef& rxRootStrm );
    void                construct( const ObjectBase& rParent );

    virtual bool        implIsValid() const;
    virtual void        implDumpHeader();
    virtual void        implDumpFooter();

    using               ObjectBase::construct;

private:
    void                constructOleStrgObj( ::oox::core::StorageRef xStrg, const ::rtl::OUString& rPath );

    void                dumpStorageInfo( bool bExtended );

private:
    ::oox::core::StorageRef mxStrg;
    ::rtl::OUString     maPath;
};

typedef ::boost::shared_ptr< OleStorageObject > OleStorageObjectRef;

// ============================================================================

class OleStorageIterator : public Base
{
public:
    explicit            OleStorageIterator( const OleStorageObject& rStrg );
    explicit            OleStorageIterator( ::oox::core::StorageRef xStrg );
                        ~OleStorageIterator();

    size_t              getElementCount() const;

    OleStorageIterator& operator++();

    ::rtl::OUString     getName() const;
    bool                isStream() const;
    bool                isStorage() const;

protected:
    void                construct( ::oox::core::StorageRef xStrg );

private:
    virtual bool        implIsValid() const;

private:
    ::oox::core::StorageRef mxStrg;
    OUStringVector      maNames;
    OUStringVector::const_iterator maIt;
};

// ============================================================================
// ============================================================================

class OleStreamObject : public BinaryStreamObject
{
public:
    explicit            OleStreamObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );
    virtual             ~OleStreamObject();

protected:
    inline explicit     OleStreamObject() {}
    void                construct( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );

    using               BinaryStreamObject::construct;
};

typedef ::boost::shared_ptr< OleStreamObject > OleStreamObjectRef;

// ============================================================================

class OlePropertyStreamObject : public OleStreamObject
{
public:
    explicit            OlePropertyStreamObject( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );

protected:
    inline explicit     OlePropertyStreamObject() {}
    void                construct( const OleStorageObject& rParentStrg, const ::rtl::OUString& rStrmName );

    virtual void        implDumpBody();

    using               OleStreamObject::construct;

private:
    void                dumpSection( const ::rtl::OUString& rGuid, sal_uInt32 nStartPos );

    void                dumpProperty( sal_Int32 nPropId, sal_uInt32 nStartPos );
    void                dumpCodePageProperty( sal_uInt32 nStartPos );
    void                dumpDictionaryProperty( sal_uInt32 nStartPos );

    void                dumpPropertyContents( sal_Int32 nPropId );
    void                dumpPropertyValue( sal_Int32 nPropId, sal_Int32 nBaseType );

    sal_Int32           dumpPropertyType();
    void                dumpBlob( const sal_Char* pcName );
    ::rtl::OUString     dumpString8( const sal_Char* pcName );
    ::rtl::OUString     dumpCharArray8( const sal_Char* pcName, sal_Int32 nCharCount );
    ::rtl::OUString     dumpString16( const sal_Char* pcName );
    ::rtl::OUString     dumpCharArray16( const sal_Char* pcName, sal_Int32 nCharCount );
    ::com::sun::star::util::DateTime dumpFileTime( const sal_Char* pcName );

    bool                startElement( sal_uInt32 nStartPos );
    void                writeSectionHeader( const ::rtl::OUString& rGuid, sal_uInt32 nStartPos );
    void                writePropertyHeader( sal_Int32 nPropId, sal_uInt32 nStartPos );

private:
    NameListRef         mxPropIds;
    rtl_TextEncoding    meTextEnc;
    bool                mbIsUnicode;
};

// ============================================================================
// ============================================================================

} // namespace dump
} // namespace oox

#endif
#endif

