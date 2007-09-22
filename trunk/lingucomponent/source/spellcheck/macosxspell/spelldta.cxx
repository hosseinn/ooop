/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: spelldta.cxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:05:54 $
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

#ifndef _COM_SUN_STAR_UNO_REFERENCE_H_
#include <com/sun/star/uno/Reference.h>
#endif

#include <com/sun/star/linguistic2/SpellFailure.hpp>
#include <com/sun/star/linguistic2/XSearchableDictionaryList.hpp>

#ifndef _TOOLS_DEBUG_HXX //autogen wg. DBG_ASSERT
#include <tools/debug.hxx>
#endif
#ifndef _UNOTOOLS_PROCESSFACTORY_HXX_
#include <unotools/processfactory.hxx>
#endif
#ifndef _OSL_MUTEX_HXX_
#include <osl/mutex.hxx>
#endif

#include "spelldta.hxx"
#include "lngsvcmgr.hxx"


using namespace utl;
using namespace osl;
using namespace rtl;
using namespace com::sun::star;
using namespace com::sun::star::beans;
using namespace com::sun::star::lang;
using namespace com::sun::star::uno;
using namespace com::sun::star::linguistic2;

namespace linguistic
{
	
///////////////////////////////////////////////////////////////////////////


#define MAX_PROPOSALS	40

Reference< XSpellAlternatives > MergeProposals(
			Reference< XSpellAlternatives > &rxAlt1,
			Reference< XSpellAlternatives > &rxAlt2)
{
	Reference< XSpellAlternatives > xMerged;

	if (!rxAlt1.is())
		xMerged = rxAlt2;
	else if (!rxAlt2.is())
		xMerged = rxAlt1;
	else
	{
		INT32 nAltCount1 = rxAlt1->getAlternativesCount();
		Sequence< OUString > aAlt1( rxAlt1->getAlternatives() );
		const OUString *pAlt1 = aAlt1.getConstArray();

		INT32 nAltCount2 = rxAlt2->getAlternativesCount();
		Sequence< OUString > aAlt2( rxAlt2->getAlternatives() );
		const OUString *pAlt2 = aAlt2.getConstArray();

		INT32 nCountNew = Min( nAltCount1 + nAltCount2, (INT32) MAX_PROPOSALS );
		Sequence< OUString > aAltNew( nCountNew );
		OUString *pAltNew = aAltNew.getArray();

		INT32 nIndex = 0;
		INT32 i = 0;
		for (int j = 0;  j < 2;  j++)
		{
			INT32 			nCount 	= j == 0 ? nAltCount1 : nAltCount2;
			const OUString  *pAlt 	= j == 0 ? pAlt1 : pAlt2;
			for (i = 0;  i < nCount  &&  nIndex < MAX_PROPOSALS;  i++)
			{
				if (pAlt[i].getLength())
					pAltNew[ nIndex++ ] = pAlt[ i ];
			}
		}
		DBG_ASSERT(nIndex == nCountNew, "lng : wrong number of proposals");

		SpellAlternatives *pSpellAlt = new SpellAlternatives;
		pSpellAlt->SetWordLanguage( rxAlt1->getWord(),
							LocaleToLanguage( rxAlt1->getLocale() ) );
		pSpellAlt->SetFailureType( rxAlt1->getFailureType() );
		pSpellAlt->SetAlternatives( aAltNew );
		xMerged = pSpellAlt;
	}

	return xMerged;
}


BOOL SeqHasEntry( 
        const Sequence< OUString > &rSeq, 
        const OUString &rTxt)
{
    BOOL bRes = FALSE;
    INT32 nLen = rSeq.getLength();
    const OUString *pEntry = rSeq.getConstArray();
    for (INT32 i = 0;  i < nLen  &&  !bRes;  ++i)
    {
        if (rTxt == pEntry[i])
            bRes = TRUE;
    }
    return bRes;
}


void SeqRemoveNegEntries( Sequence< OUString > &rSeq, 
        Reference< XDictionaryList > &rxDicList, 
        INT16 nLanguage )
{
    static const OUString aEmpty;
    BOOL bSthRemoved = FALSE;
    INT32 nLen = rSeq.getLength();
    OUString *pEntries = rSeq.getArray();
    for (INT32 i = 0;  i < nLen;  ++i)
    {
        Reference< XDictionaryEntry > xNegEntry( SearchDicList( rxDicList, 
                    pEntries[i], nLanguage, FALSE, TRUE ) );
        if (xNegEntry.is())
        {
            pEntries[i] = aEmpty;
            bSthRemoved = TRUE;
        }
    }
    if (bSthRemoved)
    {
        Sequence< OUString > aNew;
        // merge sequence without duplicates and empty strings in new empty sequence
        aNew = MergeProposalSeqs( aNew, rSeq, FALSE );
        rSeq = aNew;
    }
}


Sequence< OUString > MergeProposalSeqs(
            Sequence< OUString > &rAlt1,
            Sequence< OUString > &rAlt2,
            BOOL bAllowDuplicates )
{
    Sequence< OUString > aMerged;

    if (0 == rAlt1.getLength() && bAllowDuplicates)
        aMerged = rAlt2;
    else if (0 == rAlt2.getLength() && bAllowDuplicates)
        aMerged = rAlt1;
    else
    {
        INT32 nAltCount1 = rAlt1.getLength();
        const OUString *pAlt1 = rAlt1.getConstArray();
        INT32 nAltCount2 = rAlt2.getLength();
        const OUString *pAlt2 = rAlt2.getConstArray();

        INT32 nCountNew = Min( nAltCount1 + nAltCount2, (INT32) MAX_PROPOSALS );
        aMerged.realloc( nCountNew );
        OUString *pMerged = aMerged.getArray();

        INT32 nIndex = 0;
        INT32 i = 0;
        for (int j = 0;  j < 2;  j++)
        {
            INT32           nCount  = j == 0 ? nAltCount1 : nAltCount2;
            const OUString  *pAlt   = j == 0 ? pAlt1 : pAlt2;
            for (i = 0;  i < nCount  &&  nIndex < MAX_PROPOSALS;  i++)
            {
                if (pAlt[i].getLength() && 
                    (bAllowDuplicates || !SeqHasEntry(aMerged, pAlt[i] )))
                    pMerged[ nIndex++ ] = pAlt[ i ];
            }
        }
        //DBG_ASSERT(nIndex == nCountNew, "wrong number of proposals");
        aMerged.realloc( nIndex );
    }

    return aMerged;
}

///////////////////////////////////////////////////////////////////////////


SpellAlternatives::SpellAlternatives()
{
	nLanguage	= LANGUAGE_NONE;
	nType 		= SpellFailure::IS_NEGATIVE_WORD;
}


SpellAlternatives::SpellAlternatives( 
			const OUString &rWord, INT16 nLang,
			INT16 nFailureType, const OUString &rRplcWord ) :
	aWord		(rWord),
	nLanguage	(nLang),
	nType		(nFailureType),
	aAlt		( Sequence< OUString >(1) )
{
	if (rRplcWord.getLength())
		aAlt.getArray()[ 0 ] = rRplcWord;
	else
		aAlt.realloc( 0 );
}


SpellAlternatives::SpellAlternatives(
        const OUString &rWord, INT16 nLang, INT16 nFailureType,
        const Sequence< OUString > &rAlternatives ) :
    aWord       (rWord),
    nLanguage   (nLang),
    nType       (nFailureType),
    aAlt        (rAlternatives)
{
}


SpellAlternatives::~SpellAlternatives()
{
}


OUString SAL_CALL SpellAlternatives::getWord()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return aWord;
}


Locale SAL_CALL SpellAlternatives::getLocale()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return CreateLocale( nLanguage );
}


sal_Int16 SAL_CALL SpellAlternatives::getFailureType()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return nType;
}


sal_Int16 SAL_CALL SpellAlternatives::getAlternativesCount()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return (INT16) aAlt.getLength();
}


Sequence< OUString > SAL_CALL SpellAlternatives::getAlternatives()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return aAlt;
}


void SpellAlternatives::SetWordLanguage(const OUString &rWord, INT16 nLang)
{
	MutexGuard	aGuard( GetLinguMutex() );
	aWord = rWord;
	nLanguage = nLang;
}


void SpellAlternatives::SetFailureType(INT16 nTypeP)
{
	MutexGuard	aGuard( GetLinguMutex() );
	nType = nTypeP;
}


void SpellAlternatives::SetAlternatives( const Sequence< OUString > &rAlt )
{
	MutexGuard	aGuard( GetLinguMutex() );
	aAlt = rAlt;
}


///////////////////////////////////////////////////////////////////////////

}	// namespace linguistic

