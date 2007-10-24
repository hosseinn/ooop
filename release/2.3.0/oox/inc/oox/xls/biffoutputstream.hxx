/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: biffoutputstream.hxx,v $
 *
 *  $Revision: 1.1.2.1 $
 *
 *  last change: $Author: dr $ $Date: 2007/04/03 13:59:33 $
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

#ifndef OOX_XLS_BIFFOUTPUTSTREAM_HXX
#define OOX_XLS_BIFFOUTPUTSTREAM_HXX

#include <com/sun/star/uno/Reference.hxx>
#include "oox/xls/biffhelper.hxx"

namespace com { namespace sun { namespace star {
    namespace io { class XOutputStream; }
} } }

namespace oox {
namespace xls {

// ============================================================================

/** This class is used to export BIFF record streams.

    An instance is constructed with a com.sun.star.io.XOutputStream and the
    maximum size of BIFF record contents (e.g. in BIFF2-BIFF5: 2080 bytes, in
    BIFF8: 8224 bytes).

    To start writing a record, call startRecord(). Parameters are the record
    identifier and any calculated record size. This is for optimizing the write
    process: if the real written data has the same size as the calculated, the
    stream will not seek back and update the record size field. But it is not
    mandatory to calculate a size. Each record must be closed by calling
    endRecord(). This will check (and update) the record size field.

    If some data exceeds the record size limit, a CONTINUE record is started
    automatically and the new data will be written to this record.

    If specific data pieces must not be split, use setSliceSize(), e.g.: To
    write a sequence of 16-bit values, where 4 values form a unit and cannot be
    split, call setSliceSize(8) first (4*2 bytes == 8).

    To write unicode character arrays, call writeUnicodeBuffer(). It creates
    CONTINUE records and repeats the unicode string flag byte automatically.
*/
class BiffOutputStream
{
public:
    typedef ::com::sun::star::uno::Reference< ::com::sun::star::io::XOutputStream > XOutputStreamRef;

public:
    explicit            BiffOutputStream(
                            const XOutputStreamRef& rxOutStream );

                        ~BiffOutputStream();

    /** Sets size of data slices. 0 = no slices. */
    void                setSliceSize( sal_uInt16 nSize );

    // stream write access ----------------------------------------------------

    /** Writes nBytes bytes from the existing buffer pData.

        @return
            Number of bytes really written.
     */
    sal_uInt32          write( const void* pData, sal_uInt32 nBytes );

    BiffOutputStream&   operator<<( sal_Int8 nValue );
    BiffOutputStream&   operator<<( sal_uInt8 nValue );
    BiffOutputStream&   operator<<( sal_Int16 nValue );
    BiffOutputStream&   operator<<( sal_uInt16 nValue );
    BiffOutputStream&   operator<<( sal_Int32 nValue );
    BiffOutputStream&   operator<<( sal_uInt32 nValue );
    BiffOutputStream&   operator<<( sal_Int64 nValue );
    BiffOutputStream&   operator<<( sal_uInt64 nValue );
    BiffOutputStream&   operator<<( float fValue );
    BiffOutputStream&   operator<<( double fValue );
};

// ============================================================================

} // namespace xls
} // namespace oox

#endif

