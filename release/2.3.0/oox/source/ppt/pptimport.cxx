/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: pptimport.cxx,v $
 *
 *  $Revision: 1.1.2.9 $
 *
 *  last change: $Author: sj $ $Date: 2007/09/04 17:06:17 $
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

#include "oox/ppt/pptimport.hxx"

using ::rtl::OUString;
using namespace ::com::sun::star;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::xml::sax;
using namespace oox::core;

#define C2U(x) OUString( RTL_CONSTASCII_USTRINGPARAM( x ) )

namespace oox { namespace ppt {

OUString SAL_CALL PowerPointImport_getImplementationName() throw()
{
    return OUString( RTL_CONSTASCII_USTRINGPARAM( "com.sun.star.comp.Impress.oox.PowerPointImport" ) );
}

uno::Sequence< OUString > SAL_CALL PowerPointImport_getSupportedServiceNames() throw()
{
    const OUString aServiceName( RTL_CONSTASCII_USTRINGPARAM( "com.sun.star.comp.ooxpptx" ) );
    const uno::Sequence< OUString > aSeq( &aServiceName, 1 );
    return aSeq;
}

uno::Reference< uno::XInterface > SAL_CALL PowerPointImport_createInstance(const uno::Reference< lang::XMultiServiceFactory > & rSMgr ) throw( uno::Exception )
{
    return (cppu::OWeakObject*)new PowerPointImport( rSMgr );
}

PowerPointImport::PowerPointImport( const uno::Reference< lang::XMultiServiceFactory > & rSMgr  )
    : XmlFilterBase( rSMgr )
{
}

PowerPointImport::~PowerPointImport()
{
}

bool PowerPointImport::importDocument() throw()
{
	const OUString aEmpty;
    RelationsRef xRootRelations( getRelations( aEmpty ) );

	RelationPtr xDocRelation( xRootRelations->getRelationByType( C2U( "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ) ) );
	if( xDocRelation.get() )
	{
		Reference< XFastDocumentHandler > xParser( new PresentationFragmentHandler( this, xDocRelation->msTarget ) );
        return importFragment( xParser, xDocRelation->msTarget );
	}

	return false;
}

bool PowerPointImport::exportDocument() throw()
{
	return false;
}

sal_Int32 PowerPointImport::getSchemeClr( sal_Int32 nColorSchemeToken )
{
	sal_Int32 nColor = 0;
	if ( mpActualSlidePersist )
	{
		oox::drawingml::ClrSchemePtr pClrSchemePtr( mpActualSlidePersist->getClrScheme() );
		if ( pClrSchemePtr )
			pClrSchemePtr->getColor( nColorSchemeToken, nColor );
		else
		{
			drawingml::ThemePtr pTheme = mpActualSlidePersist->getTheme();
			if( pTheme )
			{
				pTheme->getClrScheme()->getColor( nColorSchemeToken, nColor );
			}
			else 
			{
				OSL_TRACE("OOX: PowerPointImport::mpThemePtr is NULL");
			}
		}
	}
	return nColor;
}

const oox::vml::DrawingPtr PowerPointImport::getDrawings()
{
	oox::vml::DrawingPtr xRet;
	return xRet;
}

OUString PowerPointImport::implGetImplementationName() const
{
    return PowerPointImport_getImplementationName();
}

}}
