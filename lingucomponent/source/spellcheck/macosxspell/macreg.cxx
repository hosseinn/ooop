/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: macreg.cxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:05:16 $
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

// MARKER(update_precomp.py): autogen include statement, do not remove
#include "precompiled_lingucomponent.hxx"


#include <cppuhelper/factory.hxx>	// helper for factories
#ifndef _RTL_STRING_HXX_
#include <rtl/string.hxx>
#endif

#include <com/sun/star/registry/XRegistryKey.hpp>

using namespace rtl;
using namespace com::sun::star::lang;
using namespace com::sun::star::registry;

////////////////////////////////////////
// declaration of external RegEntry-functions defined by the service objects
//

extern sal_Bool SAL_CALL MacSpellChecker_writeInfo(
	void * /*pServiceManager*/, XRegistryKey * pRegistryKey );

extern void * SAL_CALL MacSpellChecker_getFactory(
	const sal_Char * pImplName,
	XMultiServiceFactory * pServiceManager,
	void * /*pRegistryKey*/ );

////////////////////////////////////////
// definition of the two functions that are used to provide the services
//

extern "C"
{

sal_Bool SAL_CALL component_writeInfo(
	void * pServiceManager, XRegistryKey * pRegistryKey )
{
	return MacSpellChecker_writeInfo( pServiceManager, pRegistryKey );
}

void SAL_CALL component_getImplementationEnvironment(
	const sal_Char ** ppEnvTypeName, uno_Environment ** /*ppEnv*/ )
{
	*ppEnvTypeName = CPPU_CURRENT_LANGUAGE_BINDING_NAME;
}

void * SAL_CALL component_getFactory(
	const sal_Char * pImplName, void * pServiceManager, void * pRegistryKey )
{
    void * pRet = NULL;
	pRet = MacSpellChecker_getFactory(
    	pImplName,
    	reinterpret_cast< XMultiServiceFactory * >( pServiceManager ),
    	pRegistryKey );

	return pRet;
}

}

///////////////////////////////////////////////////////////////////////////

