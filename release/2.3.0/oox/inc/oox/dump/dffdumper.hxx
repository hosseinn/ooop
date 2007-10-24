/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: dffdumper.hxx,v $
 *
 *  $Revision: 1.1.2.2 $
 *
 *  last change: $Author: dr $ $Date: 2007/05/23 11:29:34 $
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

#ifndef OOX_DUMP_DFFDUMPER_HXX
#define OOX_DUMP_DFFDUMPER_HXX

#include "oox/dump/dumperbase.hxx"

#if OOX_INCLUDE_DUMPER

namespace oox {
namespace dump {

// ============================================================================
// ============================================================================

class DffRecordHeaderObject : public RecordHeaderBase
{
public:
    explicit            DffRecordHeaderObject( const InputObjectBase& rParent );

    inline sal_uInt16   getRecId() const { return mnRecId; }
    inline sal_uInt32   getRecSize() const { return mnRecSize; }
    inline sal_Int64    getBodyStart() const { return mnBodyStart; }
    inline sal_Int64    getBodyEnd() const { return mnBodyEnd; }

    inline sal_uInt16   getVer() const { return mnVer; }
    inline sal_uInt16   getInst() const { return mnInst; }

    inline bool         hasRecName() const { return getRecNames()->hasName( mnRecId ); }

protected:
    virtual bool        implIsValid() const;
    virtual void        implDumpBody();

private:
    NameListRef         mxRecInst;
    sal_Int64           mnBodyStart;
    sal_Int64           mnBodyEnd;
    sal_uInt32          mnRecSize;
    sal_uInt16          mnRecId;
    sal_uInt16          mnVer;
    sal_uInt16          mnInst;
};

// ============================================================================
// ============================================================================

class DffDumpObject : public InputObjectBase
{
public:
    explicit            DffDumpObject( const InputObjectBase& rParent );
    virtual             ~DffDumpObject();

    void                dumpDffClientPos( const sal_Char* pcName, sal_Int32 nSubScale );
    void                dumpDffClientRect();

protected:
    virtual bool        implIsValid() const;
    virtual void        implDumpBody();

private:
    void                constructDffDumpObj();

    void                dumpRecordBody();

    void                dumpDffOptRec();
    sal_uInt16          dumpDffOptPropHeader();
    void                dumpDffOptPropValue( sal_uInt16 nPropId, sal_uInt32 nValue );

private:
    typedef ::boost::shared_ptr< DffRecordHeaderObject > DffRecHeaderObjRef;
    DffRecHeaderObjRef  mxHdrObj;
};

typedef ::boost::shared_ptr< DffDumpObject > DffDumpObjectRef;

// ============================================================================
// ============================================================================

} // namespace dump
} // namespace oox

#endif
#endif

