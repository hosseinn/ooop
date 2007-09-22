/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: spelldta.hxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:06:04 $
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

#ifndef _LINGUISTIC_SPELLDTA_HXX_
#define _LINGUISTIC_SPELLDTA_HXX_


#ifndef _COM_SUN_STAR_LINGUISTIC2_XSPELLALTERNATIVES_HPP_
#include <com/sun/star/linguistic2/XSpellAlternatives.hpp>
#endif

#include <tools/solar.h>

#include <uno/lbnames.h>			// CPPU_CURRENT_LANGUAGE_BINDING_NAME macro, which specify the environment type
#include <cppuhelper/implbase1.hxx>	// helper for implementations

namespace com { namespace sun { namespace star {
    namespace linguistic2 {
        class XDictionaryList;
    }
} } }


namespace linguistic
{

///////////////////////////////////////////////////////////////////////////

::com::sun::star::uno::Reference< 
	::com::sun::star::linguistic2::XSpellAlternatives > 
		MergeProposals(
                ::com::sun::star::uno::Reference< 
                    ::com::sun::star::linguistic2::XSpellAlternatives > &rxAlt1,
                ::com::sun::star::uno::Reference< 
                    ::com::sun::star::linguistic2::XSpellAlternatives > &rxAlt2 );

::com::sun::star::uno::Sequence< ::rtl::OUString > 
        MergeProposalSeqs(
                ::com::sun::star::uno::Sequence< ::rtl::OUString > &rAlt1,
                ::com::sun::star::uno::Sequence< ::rtl::OUString > &rAlt2,
                BOOL bAllowDuplicates );

void    SeqRemoveNegEntries( 
                ::com::sun::star::uno::Sequence< ::rtl::OUString > &rSeq,
                ::com::sun::star::uno::Reference< 
                    ::com::sun::star::linguistic2::XDictionaryList > &rxDicList, 
                INT16 nLanguage );

BOOL    SeqHasEntry( 
                const ::com::sun::star::uno::Sequence< ::rtl::OUString > &rSeq, 
                const ::rtl::OUString &rTxt);

///////////////////////////////////////////////////////////////////////////


class SpellAlternatives :
	public cppu::WeakImplHelper1
	<
		::com::sun::star::linguistic2::XSpellAlternatives
	>
{
	::com::sun::star::uno::Sequence< ::rtl::OUString >	aAlt;	// list of alternatives, may be empty.
	::rtl::OUString			aWord;
	INT16					nType;			// type of failure
	INT16					nLanguage;

	// disallow copy-constructor and assignment-operator for now
	SpellAlternatives(const SpellAlternatives &);
	SpellAlternatives & operator = (const SpellAlternatives &);

public:
	SpellAlternatives();
	SpellAlternatives(const ::rtl::OUString &rWord, INT16 nLang, INT16 nFailureType,
					  const ::rtl::OUString &rRplcWord );
    SpellAlternatives(const ::rtl::OUString &rWord, INT16 nLang, INT16 nFailureType,
                      const ::com::sun::star::uno::Sequence< ::rtl::OUString > &rAlternatives );
	virtual ~SpellAlternatives();

	// XSpellAlternatives
    virtual ::rtl::OUString SAL_CALL
		getWord()
			throw(::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::lang::Locale SAL_CALL
		getLocale()
			throw(::com::sun::star::uno::RuntimeException);
    virtual sal_Int16 SAL_CALL
		getFailureType()
			throw(::com::sun::star::uno::RuntimeException);
    virtual sal_Int16 SAL_CALL
		getAlternativesCount()
			throw(::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Sequence< ::rtl::OUString > SAL_CALL
		getAlternatives()
			throw(::com::sun::star::uno::RuntimeException);

	// non-interface specific functions
	void	SetWordLanguage(const ::rtl::OUString &rWord, INT16 nLang);
	void	SetFailureType(INT16 nTypeP);
	void	SetAlternatives(
				const ::com::sun::star::uno::Sequence< ::rtl::OUString > &rAlt );
};


///////////////////////////////////////////////////////////////////////////

}	// namespace linguistic

#endif

