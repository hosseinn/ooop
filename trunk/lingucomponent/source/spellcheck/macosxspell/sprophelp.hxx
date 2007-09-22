/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: sprophelp.hxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:06:21 $
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

#ifndef _LINGU2_PROPHELP_HXX_
#define _LINGU2_PROPHELP_HXX_

#include <tools/solar.h>

#include <uno/lbnames.h>			// CPPU_CURRENT_LANGUAGE_BINDING_NAME macro, which specify the environment type
#include <cppuhelper/implbase2.hxx>	// helper for implementations

#ifndef _CPPUHELPER_INTERFACECONTAINER_H_
#include <cppuhelper/interfacecontainer.h>
#endif

#ifndef _COM_SUN_STAR_BEANS_XPROPERTYCHANGELISTENER_HPP_
#include <com/sun/star/beans/XPropertyChangeListener.hpp>
#endif
#ifndef _COM_SUN_STAR_BEANS_PROPERTYVALUES_HPP_
#include <com/sun/star/beans/PropertyValues.hpp>
#endif

#include <com/sun/star/linguistic2/XLinguServiceEventBroadcaster.hpp>

namespace com { namespace sun { namespace star { namespace beans {
	class	XPropertySet;
}}}}

namespace com { namespace sun { namespace star { namespace linguistic2 {
	struct	LinguServiceEvent;
}}}}


using namespace ::rtl;
using namespace ::com::sun::star::uno;
using namespace ::com::sun::star::beans;
using namespace ::com::sun::star::lang;
using namespace ::com::sun::star::linguistic2;

///////////////////////////////////////////////////////////////////////////
// PropertyChgHelper
// virtual base class for all XPropertyChangeListener members of the
// various lingu services.
// Only propertyChange needs to be implemented.

class PropertyChgHelper :
	public cppu::WeakImplHelper2
	<
		XPropertyChangeListener,
		XLinguServiceEventBroadcaster
	>
{
	Sequence< OUString >  	   			aPropNames;
	Reference< XInterface >	 			xMyEvtObj;
	::cppu::OInterfaceContainerHelper	aLngSvcEvtListeners;
	Reference< XPropertySet >  			xPropSet;

	// disallow use of copy-constructor and assignment-operator
	PropertyChgHelper( const PropertyChgHelper & );
	PropertyChgHelper & operator = ( const PropertyChgHelper & );

public:
	PropertyChgHelper( 
			const Reference< XInterface > &rxSource,
			Reference< XPropertySet > &rxPropSet,
			const char *pPropNames[], USHORT nPropCount );
	virtual ~PropertyChgHelper();

	// XEventListener
	virtual void SAL_CALL 
		disposing( const EventObject& rSource ) 
			throw(RuntimeException);

	// XPropertyChangeListener
	virtual void SAL_CALL 
		propertyChange( const PropertyChangeEvent& rEvt ) 
			throw(RuntimeException) = 0;

	// XLinguServiceEventBroadcaster
    virtual sal_Bool SAL_CALL 
		addLinguServiceEventListener( 
				const Reference< XLinguServiceEventListener >& rxListener ) 
			throw(RuntimeException);
    virtual sal_Bool SAL_CALL 
		removeLinguServiceEventListener( 
				const Reference< XLinguServiceEventListener >& rxListener ) 
			throw(RuntimeException);

	// non UNO functions
	void 	AddAsPropListener();
	void	RemoveAsPropListener();
	void	LaunchEvent( const LinguServiceEvent& rEvt );
	
	const Sequence< OUString > &
			GetPropNames() const	{ return aPropNames; }
	const Reference< XPropertySet >	&
			GetPropSet() const		{ return xPropSet; }
	const Reference< XInterface > &
			GetEvtObj() const		{ return xMyEvtObj; }
};


///////////////////////////////////////////////////////////////////////////


class PropertyHelper_Spell :
	public PropertyChgHelper
{
	// default values
	BOOL	bIsGermanPreReform;
	BOOL	bIsIgnoreControlCharacters;
	BOOL	bIsUseDictionaryList;
	BOOL	bIsSpellUpperCase;
	BOOL	bIsSpellWithDigits;
	BOOL	bIsSpellCapitalization;

	// return values, will be set to default value or current temporary value
	BOOL	bResIsGermanPreReform;
	BOOL	bResIsIgnoreControlCharacters;
	BOOL	bResIsUseDictionaryList;
	BOOL	bResIsSpellUpperCase;
	BOOL	bResIsSpellWithDigits;
	BOOL	bResIsSpellCapitalization;

	
	// disallow use of copy-constructor and assignment-operator
	PropertyHelper_Spell( const PropertyHelper_Spell & );
	PropertyHelper_Spell & operator = ( const PropertyHelper_Spell & );
		
	void	SetDefault();

public:
	PropertyHelper_Spell(
			const Reference< XInterface > &rxSource,
			Reference< XPropertySet > &rxPropSet );
	virtual ~PropertyHelper_Spell();

	// XPropertyChangeListener
	virtual void SAL_CALL 
		propertyChange( const PropertyChangeEvent& rEvt ) 
			throw(RuntimeException);
	
	void	SetTmpPropVals( const PropertyValues &rPropVals );
	
	BOOL	IsGermanPreReform() const			{ return bResIsGermanPreReform; }
	BOOL	IsIgnoreControlCharacters() const	{ return bResIsIgnoreControlCharacters; }
	BOOL	IsUseDictionaryList() const			{ return bResIsUseDictionaryList; }
	BOOL	IsSpellUpperCase() const			{ return bResIsSpellUpperCase; }
	BOOL	IsSpellWithDigits() const			{ return bResIsSpellWithDigits; }
	BOOL	IsSpellCapitalization() const		{ return bResIsSpellCapitalization; }
};

///////////////////////////////////////////////////////////////////////////

#endif

