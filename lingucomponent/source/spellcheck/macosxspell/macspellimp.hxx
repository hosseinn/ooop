/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: macspellimp.hxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:05:34 $
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

#ifndef _MACSPELLIMP_H_
#define _MACSPELLIMP_H_

#include <uno/lbnames.h>			// CPPU_CURRENT_LANGUAGE_BINDING_NAME macro, which specify the environment type
#include <cppuhelper/implbase1.hxx>	// helper for implementations
#include <cppuhelper/implbase6.hxx>	// helper for implementations

#ifdef MACOSX
#include <premac.h>
#include <Carbon/Carbon.h>
#import <Cocoa/Cocoa.h>
#include <postmac.h>
#endif

#ifndef _COM_SUN_STAR_LANG_XCOMPONENT_HPP_
#include <com/sun/star/lang/XComponent.hpp>
#endif
#ifndef _COM_SUN_STAR_LANG_XINITIALIZATION_HPP_
#include <com/sun/star/lang/XInitialization.hpp>
#endif
#ifndef _COM_SUN_STAR_LANG_XSERVICEDISPLAYNAME_HPP_
#include <com/sun/star/lang/XServiceDisplayName.hpp>
#endif
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/beans/PropertyValues.hpp>
#include <com/sun/star/lang/XServiceInfo.hpp>
#include <com/sun/star/linguistic2/XSpellChecker.hpp>
#include <com/sun/star/linguistic2/XSearchableDictionaryList.hpp>
#include <com/sun/star/linguistic2/XLinguServiceEventBroadcaster.hpp>

#ifndef _TOOLS_TABLE_HXX
#include <tools/table.hxx>
#endif

#include <linguistic/misc.hxx>
#include "sprophelp.hxx"

using namespace ::rtl;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::linguistic2;


#define A2OU(x)	::rtl::OUString::createFromAscii( x )

#define OU2A(rtlOUString)     ::rtl::OString((rtlOUString).getStr(), (rtlOUString).getLength(), RTL_TEXTENCODING_ASCII_US).getStr()

#define OU2UTF8(rtlOUString)     ::rtl::OString((rtlOUString).getStr(), (rtlOUString).getLength(), RTL_TEXTENCODING_UTF8).getStr()

#define OU2ISO_1(rtlOUString)     ::rtl::OString((rtlOUString).getStr(), (rtlOUString).getLength(), RTL_TEXTENCODING_ISO_8859_1).getStr()

#define OU2ENC(rtlOUString, rtlEncoding)     ::rtl::OString((rtlOUString).getStr(), (rtlOUString).getLength(), rtlEncoding, RTL_UNICODETOTEXT_FLAGS_UNDEFINED_QUESTIONMARK).getStr()


///////////////////////////////////////////////////////////////////////////


class MacSpellChecker :
	public cppu::WeakImplHelper6
	<
		XSpellChecker,
		XLinguServiceEventBroadcaster,
		XInitialization,
		XComponent,
		XServiceInfo,
		XServiceDisplayName
	>
{
	Sequence< Locale >                 aSuppLocales;
//        Hunspell **                         aDicts;
        rtl_TextEncoding *                 aDEncs;
        Locale *                           aDLocs;
        OUString *                         aDNames;
        sal_Int32                          numdict;
		NSSpellChecker *					macSpell;
		int 								macTag;   //unique tag for this doc

	::cppu::OInterfaceContainerHelper		aEvtListeners;
	Reference< XPropertyChangeListener >	xPropHelper;
	PropertyHelper_Spell *			 		pPropHelper;
	BOOL									bDisposing;

	// disallow copy-constructor and assignment-operator for now
	MacSpellChecker(const MacSpellChecker &);
	MacSpellChecker & operator = (const MacSpellChecker &);

	PropertyHelper_Spell &	GetPropHelper_Impl();
	PropertyHelper_Spell &	GetPropHelper()
	{
		return pPropHelper ? *pPropHelper : GetPropHelper_Impl();
	}

	INT16	GetSpellFailure( const OUString &rWord, const Locale &rLocale );
	Reference< XSpellAlternatives >
			GetProposals( const OUString &rWord, const Locale &rLocale );

public:
	MacSpellChecker();
	virtual ~MacSpellChecker();

	// XSupportedLocales (for XSpellChecker)
    virtual Sequence< Locale > SAL_CALL 
		getLocales() 
			throw(RuntimeException);
    virtual sal_Bool SAL_CALL 
		hasLocale( const Locale& rLocale ) 
			throw(RuntimeException);

	// XSpellChecker
    virtual sal_Bool SAL_CALL 
		isValid( const OUString& rWord, const Locale& rLocale, 
				const PropertyValues& rProperties ) 
			throw(IllegalArgumentException, 
				  RuntimeException);
    virtual Reference< XSpellAlternatives > SAL_CALL 
		spell( const OUString& rWord, const Locale& rLocale, 
				const PropertyValues& rProperties ) 
			throw(IllegalArgumentException, 
				  RuntimeException);

    // XLinguServiceEventBroadcaster
    virtual sal_Bool SAL_CALL 
		addLinguServiceEventListener( 
			const Reference< XLinguServiceEventListener >& rxLstnr ) 
			throw(RuntimeException);
    virtual sal_Bool SAL_CALL 
		removeLinguServiceEventListener( 
			const Reference< XLinguServiceEventListener >& rxLstnr ) 
			throw(RuntimeException);
	
	// XServiceDisplayName
    virtual OUString SAL_CALL 
		getServiceDisplayName( const Locale& rLocale ) 
			throw(RuntimeException);

	// XInitialization
    virtual void SAL_CALL 
		initialize( const Sequence< Any >& rArguments ) 
			throw(Exception, RuntimeException);

	// XComponent
	virtual void SAL_CALL 
		dispose() 
			throw(RuntimeException);
    virtual void SAL_CALL 
		addEventListener( const Reference< XEventListener >& rxListener ) 
			throw(RuntimeException);
    virtual void SAL_CALL 
		removeEventListener( const Reference< XEventListener >& rxListener ) 
			throw(RuntimeException);

	////////////////////////////////////////////////////////////
	// Service specific part
	//

	// XServiceInfo
    virtual OUString SAL_CALL 
		getImplementationName() 
			throw(RuntimeException);
    virtual sal_Bool SAL_CALL 
		supportsService( const OUString& rServiceName ) 
			throw(RuntimeException);
    virtual Sequence< OUString > SAL_CALL 
		getSupportedServiceNames() 
			throw(RuntimeException);


	static inline OUString	
		getImplementationName_Static() throw();
    static Sequence< OUString >	
		getSupportedServiceNames_Static() throw();
};

inline OUString MacSpellChecker::getImplementationName_Static() throw()
{
	return A2OU( "org.openoffice.lingu.MacOSXSpellChecker" );
}



///////////////////////////////////////////////////////////////////////////

#endif

