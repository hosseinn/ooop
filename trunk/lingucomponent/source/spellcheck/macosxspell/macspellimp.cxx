/*************************************************************************
 *
 *  OpenOffice.org - a multi-platform office productivity suite
 *
 *  $RCSfile: macspellimp.cxx,v $
 *
 *  $Revision: 1.2 $
 *
 *  last change: $Author: ihi $ $Date: 2007/09/13 18:05:25 $
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
#ifndef _COM_SUN_STAR_LINGUISTIC2_XSEARCHABLEDICTIONARYLIST_HPP_
#include <com/sun/star/linguistic2/XSearchableDictionaryList.hpp>
#endif

#include <com/sun/star/linguistic2/SpellFailure.hpp>
#include <cppuhelper/factory.hxx>	// helper for factories
#include <com/sun/star/registry/XRegistryKey.hpp>

#ifndef _TOOLS_DEBUG_HXX //autogen wg. DBG_ASSERT
#include <tools/debug.hxx>
#endif
#ifndef _UNOTOOLS_PROCESSFACTORY_HXX_
#include <unotools/processfactory.hxx>
#endif
#ifndef _OSL_MUTEX_HXX_
#include <osl/mutex.hxx>
#endif

//#include <hunspell.hxx>
#include <dictmgr.hxx>

#ifndef _MACSPELLIMP_H_
#include <macspellimp.hxx>
#endif

//#include <linguistic/lngprops.hxx>
#include "spelldta.hxx"

#ifndef INCLUDED_SVTOOLS_PATHOPTIONS_HXX
#include <svtools/pathoptions.hxx>
#endif
#ifndef INCLUDED_SVTOOLS_USEROPTIONS_HXX
#include <svtools/useroptions.hxx>
#endif
#include <osl/file.hxx>
#include <rtl/ustrbuf.hxx>


using namespace utl;
using namespace osl;
using namespace rtl;
using namespace com::sun::star;
using namespace com::sun::star::beans;
using namespace com::sun::star::lang;
using namespace com::sun::star::uno;
using namespace com::sun::star::linguistic2;
using namespace linguistic;
///////////////////////////////////////////////////////////////////////////
// dbg_dump for development
#if OSL_DEBUG_LEVEL > 1
#include <rtl/strbuf.hxx>
#include <rtl/ustring.hxx>

const sal_Char *dbg_dump(const rtl::OString &rStr)
{
    static rtl::OStringBuffer aStr;

    aStr = rtl::OStringBuffer(rStr);
    aStr.append(static_cast<char>(0));
    return aStr.getStr();
}

const sal_Char *dbg_dump(const rtl::OUString &rStr)
{
    return dbg_dump(rtl::OUStringToOString(rStr, RTL_TEXTENCODING_UTF8));
}

const sal_Char *dbg_dump(rtl_String *pStr)
{
    return dbg_dump(rtl::OString(pStr));
}

const sal_Char *dbg_dump(rtl_uString *pStr)
{
    return dbg_dump(rtl::OUString(pStr));
}

#endif
///////////////////////////////////////////////////////////////////////////

static BOOL operator == ( const Locale &rL1, const Locale &rL2 )
{
	return	rL1.Language ==  rL2.Language	&&
			rL1.Country  ==  rL2.Country	&&
			rL1.Variant  ==  rL2.Variant;
}

///////////////////////////////////////////////////////////////////////////

MacSpellChecker::MacSpellChecker() :
	aEvtListeners	( GetLinguMutex() )
{
//    aDicts = NULL;
	aDEncs = NULL;
	aDLocs = NULL;
	aDNames = NULL;
	bDisposing = FALSE;
	pPropHelper = NULL;
    numdict = 0;
    NSApplicationLoad();
	NSAutoreleasePool* pool	= [[NSAutoreleasePool alloc] init];
    macSpell = [NSSpellChecker sharedSpellChecker];
    macTag = [NSSpellChecker uniqueSpellDocumentTag];
    [pool release];
}


MacSpellChecker::~MacSpellChecker()
{
  // if (aDicts) {
  //    for (int i = 0; i < numdict; i++) { 
  //           if (aDicts[i]) delete aDicts[i];
  //           aDicts[i] = NULL;
  //    }
  //    delete[] aDicts;
  // }
  // aDicts = NULL;
  numdict = 0;	
  if (aDEncs) delete[] aDEncs;
  aDEncs = NULL;
  if (aDLocs) delete[] aDLocs;
  aDLocs = NULL;
  if (aDNames) delete[] aDNames;
  aDNames = NULL;
  if (pPropHelper)
	 pPropHelper->RemoveAsPropListener();
}


PropertyHelper_Spell & MacSpellChecker::GetPropHelper_Impl()
{
	if (!pPropHelper)
	{
		Reference< XPropertySet	>	xPropSet( GetLinguProperties(), UNO_QUERY );

		pPropHelper	= new PropertyHelper_Spell( (XSpellChecker *) this, xPropSet );
		xPropHelper = pPropHelper;
		pPropHelper->AddAsPropListener();	//! after a reference is established
	}
	return *pPropHelper;
}


Sequence< Locale > SAL_CALL MacSpellChecker::getLocales()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

        // this routine should return the locales supported by the installed
        // dictionaries.  So here we need to parse both the user edited 
        // dictionary list and the shared dictionary list 
        // to see what dictionaries the admin/user has installed

        int numusr;          // number of user dictionary entries
        int numshr;          // number of shared dictionary entries
        dictentry * spdict;  // shared dict entry pointer
        dictentry * updict;  // user dict entry pointer
        SvtPathOptions aPathOpt;
		rtl_TextEncoding aEnc = RTL_TEXTENCODING_UTF8;

    	std::vector<objc_object *> postspdict;
    	//std::vector<dictentry *> postspdict;
        std::vector<dictentry *> postupdict;


	if (!numdict) {

        // invoke a dictionary manager to get the user dictionary list
        // TODO How on Mac OS X?

        // invoke a second  dictionary manager to get the shared dictionary list
        NSArray *aLocales = [NSLocale availableLocaleIdentifiers];

        //Test for existence of the dictionaries
        for (unsigned int i = 0; i < [aLocales count]; i++)
        {
            if( [macSpell setLanguage:[aLocales objectAtIndex:i] ] )
            {
                postspdict.push_back( [ aLocales objectAtIndex:i ] );
            }
        }

	    numusr = postupdict.size();
        numshr = postspdict.size();

        // we really should merge these and remove duplicates but since 
        // users can name their dictionaries anything they want it would
        // be impossible to know if a real duplication exists unless we
        // add some unique key to each myspell dictionary
        numdict = numshr + numusr;

        if (numdict) {
            aDLocs = new Locale [numdict];
            aDEncs  = new rtl_TextEncoding [numdict];
            aDNames = new OUString [numdict];
            aSuppLocales.realloc(numdict);
            Locale * pLocale = aSuppLocales.getArray();
            int numlocs = 0;
            int newloc;
            int i,j;
            int k = 0;
            
            //first add the user dictionaries
            //TODO for MAC?
            
            // now add the shared dictionaries
            for (i = 0; i < numshr; i++) {
                NSDictionary *aLocDict = [ NSLocale componentsFromLocaleIdentifier:postspdict[i] ];
                NSString* aLang = [ aLocDict objectForKey:NSLocaleLanguageCode ];
                NSString* aCountry = [ aLocDict objectForKey:NSLocaleCountryCode ];
                OUString lang([aLang cStringUsingEncoding: NSUTF8StringEncoding], [aLang length], aEnc); 
                OUString country([ aCountry cStringUsingEncoding: NSUTF8StringEncoding], [aCountry length], aEnc); 
                Locale nLoc( lang, country, OUString() );
                newloc = 1;
                //eliminate duplicates (is this needed for MacOS?)
                for (j = 0; j < numlocs; j++) {
                    if (nLoc == pLocale[j]) newloc = 0;
                }
                if (newloc) {
                    pLocale[numlocs] = nLoc;
                    numlocs++;
                }
                aDLocs[k] = nLoc;
                //pointer to Hunspell dictionary - not needed for MAC
                //aDicts[k] = NULL;
                aDEncs[k] = 0;
                // Dictionary file names not valid for Mac Spell
                //aDNames[k] = aPathOpt.GetLinguisticPath() + A2OU("/ooo/") + A2OU(postspdict[i]->filename);
                k++;
            }
                            
            aSuppLocales.realloc(numlocs);
                            
        } else {
            /* no dictionary.lst found so register no dictionaries */
            numdict = 0;
            //aDicts = NULL;
                aDEncs  = NULL;
                aDLocs = NULL;
                aDNames = NULL;
	            aSuppLocales.realloc(0);
            }

            /* de-allocation of memory is handled inside the DictMgr */
            updict = NULL;
            spdict = NULL;

        }

	return aSuppLocales;
}



sal_Bool SAL_CALL MacSpellChecker::hasLocale(const Locale& rLocale)
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	BOOL bRes = FALSE;
	if (!aSuppLocales.getLength())
		getLocales();

	INT32 nLen = aSuppLocales.getLength();
	for (INT32 i = 0;  i < nLen;  ++i)
	{
		const Locale *pLocale = aSuppLocales.getConstArray();
		if (rLocale == pLocale[i])
		{
			bRes = TRUE;
			break;
		}
	}
	return bRes;
}


INT16 MacSpellChecker::GetSpellFailure( const OUString &rWord, const Locale &rLocale )
{       
    rtl_TextEncoding aEnc;

	// initialize a myspell object for each dictionary once 
        // (note: mutex is held higher up in isValid)


	INT16 nRes = -1;

        // first handle smart quotes both single and double
	OUStringBuffer rBuf(rWord);
        sal_Int32 n = rBuf.getLength();
        sal_Unicode c;
	for (sal_Int32 ix=0; ix < n; ix++) {
	    c = rBuf.charAt(ix);
            if ((c == 0x201C) || (c == 0x201D)) rBuf.setCharAt(ix,(sal_Unicode)0x0022);
            if ((c == 0x2018) || (c == 0x2019)) rBuf.setCharAt(ix,(sal_Unicode)0x0027);
        }
        OUString nWord(rBuf.makeStringAndClear());

	if (n)
	{
        aEnc = 0;
		NSAutoreleasePool* pool	= [[NSAutoreleasePool alloc] init];
        NSString* aNSStr = [[NSString alloc] initWithCharacters: nWord.getStr() length: nWord.getLength()];
        NSString* aLang = [[NSString alloc] initWithCharacters: rLocale.Language.getStr() length: rLocale.Language.getLength()];
        if(rLocale.Country.getLength()>0)
        {
            NSString* aCountry = [[NSString alloc] initWithCharacters: rLocale.Country.getStr() length: rLocale.Country.getLength()];
            NSString* aTag = @"_";
            NSString* aTaggedCountry = [aTag stringByAppendingString:aCountry];
            [aLang autorelease];
            aLang = [aLang  stringByAppendingString:aTaggedCountry];
        }
        
        int aCount;
        NSRange range = [macSpell checkSpellingOfString:aNSStr startingAt:0 language:aLang wrap:FALSE inSpellDocumentWithTag:macTag wordCount:&aCount];
		int rVal = 0; 
		if(range.length>0)
		{
			rVal = -1;
		}
		else
		{
			rVal = 1;
		}
        [pool release];
        if (rVal != 1)
        {  
            nRes = SpellFailure::SPELLING_ERROR;
        } else {
            return -1;
        }
    }
	return nRes;
}



sal_Bool SAL_CALL 
	MacSpellChecker::isValid( const OUString& rWord, const Locale& rLocale, 
			const PropertyValues& rProperties ) 
		throw(IllegalArgumentException, RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

 	if (rLocale == Locale()  ||  !rWord.getLength())
		return TRUE;

	if (!hasLocale( rLocale ))
#ifdef LINGU_EXCEPTIONS
		throw( IllegalArgumentException() );
#else
		return TRUE;
#endif

	// Get property values to be used.
	// These are be the default values set in the SN_LINGU_PROPERTIES
	// PropertySet which are overridden by the supplied ones from the
	// last argument.
	// You'll probably like to use a simplier solution than the provided
	// one using the PropertyHelper_Spell.

	PropertyHelper_Spell &rHelper = GetPropHelper();
	rHelper.SetTmpPropVals( rProperties );

	INT16 nFailure = GetSpellFailure( rWord, rLocale );
	if (nFailure != -1)
	{
		INT16 nLang = LocaleToLanguage( rLocale );
		// postprocess result for errors that should be ignored
		if (   (!rHelper.IsSpellUpperCase()  && IsUpper( rWord, nLang ))
			|| (!rHelper.IsSpellWithDigits() && HasDigits( rWord ))
			|| (!rHelper.IsSpellCapitalization()
				&&  nFailure == SpellFailure::CAPTION_ERROR)
		)
			nFailure = -1;
	}

	return (nFailure == -1);
}


Reference< XSpellAlternatives >
	MacSpellChecker::GetProposals( const OUString &rWord, const Locale &rLocale )
{
	// Retrieves the return values for the 'spell' function call in case
	// of a misspelled word.
	// Especially it may give a list of suggested (correct) words:
	
	Reference< XSpellAlternatives > xRes;
        // note: mutex is held by higher up by spell which covers both

    INT16 nLang = LocaleToLanguage( rLocale );
    rtl_TextEncoding aEnc = RTL_TEXTENCODING_UTF8;
	int count;
	Sequence< OUString > aStr( 0 );

        // first handle smart quotes (single and double)
	OUStringBuffer rBuf(rWord);
        sal_Int32 n = rBuf.getLength();
        sal_Unicode c;
	for (sal_Int32 ix=0; ix < n; ix++) {
	     c = rBuf.charAt(ix);
             if ((c == 0x201C) || (c == 0x201D)) rBuf.setCharAt(ix,(sal_Unicode)0x0022);
             if ((c == 0x2018) || (c == 0x2019)) rBuf.setCharAt(ix,(sal_Unicode)0x0027);
        }
        OUString nWord(rBuf.makeStringAndClear());

	if (n)
	{
		NSAutoreleasePool* pool	= [[NSAutoreleasePool alloc] init];
        NSString* aNSStr = [[NSString alloc] initWithCharacters: nWord.getStr() length: nWord.getLength()];
        NSString* aLang = [[NSString alloc] initWithCharacters: rLocale.Language.getStr() length: rLocale.Language.getLength() ];
        if(rLocale.Country.getLength()>0)
        {
            NSString* aCountry = [[NSString alloc] initWithCharacters: rLocale.Country.getStr() length: rLocale.Country.getLength() ];
            NSString* aTag = @"_";
            NSString* aTaggedCountry = [aTag stringByAppendingString:aCountry];
            [aLang autorelease];
            aLang = [aLang  stringByAppendingString:aTaggedCountry];
        }
        [macSpell setLanguage:aLang];
        NSArray *guesses = [macSpell guessesForWord:aNSStr];
        count = [guesses count];
        if (count) 
        {
           aStr.realloc( count );
           OUString *pStr = aStr.getArray();
               for (int ii=0; ii < count; ii++)  
               {  
                  // if needed add: if (suglst[ii] == NULL) continue;
                  NSString* guess = [guesses objectAtIndex:ii];
                  OUString cvtwrd([guess cStringUsingEncoding: NSUTF8StringEncoding], [guess length], aEnc);  
                  pStr[ii] = cvtwrd;  
               }  
        }
       [pool release];
	}

            // now return an empty alternative for no suggestions or the list of alternatives if some found
	    SpellAlternatives *pAlt = new SpellAlternatives;
            String aTmp(rWord);
	    pAlt->SetWordLanguage( aTmp, nLang );
	    pAlt->SetFailureType( SpellFailure::SPELLING_ERROR );
	    pAlt->SetAlternatives( aStr );
	    xRes = pAlt;
        return xRes;

}




Reference< XSpellAlternatives > SAL_CALL 
	MacSpellChecker::spell( const OUString& rWord, const Locale& rLocale, 
			const PropertyValues& rProperties ) 
		throw(IllegalArgumentException, RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

 	if (rLocale == Locale()  ||  !rWord.getLength())
		return NULL;

	if (!hasLocale( rLocale ))
#ifdef LINGU_EXCEPTIONS
		throw( IllegalArgumentException() );
#else
		return NULL;
#endif

	Reference< XSpellAlternatives > xAlt;
	if (!isValid( rWord, rLocale, rProperties ))
	{
		xAlt =  GetProposals( rWord, rLocale );
	}
	return xAlt;
}


Reference< XInterface > SAL_CALL MacSpellChecker_CreateInstance( 
            const Reference< XMultiServiceFactory > & /*rSMgr*/ )
		throw(Exception)
{

	Reference< XInterface > xService = (cppu::OWeakObject*) new MacSpellChecker;
	return xService;
}
    
	
sal_Bool SAL_CALL 
	MacSpellChecker::addLinguServiceEventListener( 
			const Reference< XLinguServiceEventListener >& rxLstnr ) 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	BOOL bRes = FALSE;   
	if (!bDisposing && rxLstnr.is())
	{
		bRes = GetPropHelper().addLinguServiceEventListener( rxLstnr );
	}
	return bRes;
}


sal_Bool SAL_CALL 
	MacSpellChecker::removeLinguServiceEventListener( 
			const Reference< XLinguServiceEventListener >& rxLstnr ) 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	BOOL bRes = FALSE;   
	if (!bDisposing && rxLstnr.is())
	{
		DBG_ASSERT( xPropHelper.is(), "xPropHelper non existent" );
		bRes = GetPropHelper().removeLinguServiceEventListener( rxLstnr );
	}
	return bRes;
}


OUString SAL_CALL 
    MacSpellChecker::getServiceDisplayName( const Locale& /*rLocale*/ ) 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	return A2OU( "Mac OS X Spell Checker" );
}


void SAL_CALL 
	MacSpellChecker::initialize( const Sequence< Any >& rArguments ) 
		throw(Exception, RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	if (!pPropHelper)
	{
		INT32 nLen = rArguments.getLength();
		if (2 == nLen)
		{
			Reference< XPropertySet	>	xPropSet;
			rArguments.getConstArray()[0] >>= xPropSet;
			//rArguments.getConstArray()[1] >>= xDicList;

			//! Pointer allows for access of the non-UNO functions.
			//! And the reference to the UNO-functions while increasing
			//! the ref-count and will implicitly free the memory
			//! when the object is not longer used.
			pPropHelper = new PropertyHelper_Spell( (XSpellChecker *) this, xPropSet );
			xPropHelper = pPropHelper;
			pPropHelper->AddAsPropListener();	//! after a reference is established
		}
		else
			DBG_ERROR( "wrong number of arguments in sequence" );

	}
}


void SAL_CALL 
	MacSpellChecker::dispose() 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	if (!bDisposing)
	{
		bDisposing = TRUE;
		EventObject	aEvtObj( (XSpellChecker *) this );
		aEvtListeners.disposeAndClear( aEvtObj );
	}
}


void SAL_CALL 
	MacSpellChecker::addEventListener( const Reference< XEventListener >& rxListener ) 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	if (!bDisposing && rxListener.is())
		aEvtListeners.addInterface( rxListener );
}


void SAL_CALL 
	MacSpellChecker::removeEventListener( const Reference< XEventListener >& rxListener ) 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );
	
	if (!bDisposing && rxListener.is())
		aEvtListeners.removeInterface( rxListener );
}


///////////////////////////////////////////////////////////////////////////
// Service specific part
//

OUString SAL_CALL MacSpellChecker::getImplementationName() 
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	return getImplementationName_Static();
}


sal_Bool SAL_CALL MacSpellChecker::supportsService( const OUString& ServiceName )
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	Sequence< OUString > aSNL = getSupportedServiceNames();
	const OUString * pArray = aSNL.getConstArray();
	for( INT32 i = 0; i < aSNL.getLength(); i++ )
		if( pArray[i] == ServiceName )
			return TRUE;
	return FALSE;
}


Sequence< OUString > SAL_CALL MacSpellChecker::getSupportedServiceNames()
		throw(RuntimeException)
{
	MutexGuard	aGuard( GetLinguMutex() );

	return getSupportedServiceNames_Static();
}


Sequence< OUString > MacSpellChecker::getSupportedServiceNames_Static() 
		throw()
{
	MutexGuard	aGuard( GetLinguMutex() );

	Sequence< OUString > aSNS( 1 );	// auch mehr als 1 Service moeglich
	aSNS.getArray()[0] = A2OU( SN_SPELLCHECKER );
	return aSNS;
}


sal_Bool SAL_CALL MacSpellChecker_writeInfo(
			void * /*pServiceManager*/, registry::XRegistryKey * pRegistryKey )
{

	try
	{
		String aImpl( '/' );
		aImpl += MacSpellChecker::getImplementationName_Static().getStr();
		aImpl.AppendAscii( "/UNO/SERVICES" );
		Reference< registry::XRegistryKey > xNewKey =
				pRegistryKey->createKey( aImpl );
		Sequence< OUString > aServices =
				MacSpellChecker::getSupportedServiceNames_Static();
		for( INT32 i = 0; i < aServices.getLength(); i++ )
			xNewKey->createKey( aServices.getConstArray()[i] );

		return sal_True;
	}
	catch(Exception &)
	{
		return sal_False;
	}
}


void * SAL_CALL MacSpellChecker_getFactory( const sal_Char * pImplName,
			XMultiServiceFactory * pServiceManager, void *  )
{
	void * pRet = 0;
	if ( !MacSpellChecker::getImplementationName_Static().compareToAscii( pImplName ) )
	{
		Reference< XSingleServiceFactory > xFactory =
			cppu::createOneInstanceFactory(
				pServiceManager,
				MacSpellChecker::getImplementationName_Static(),
				MacSpellChecker_CreateInstance,
				MacSpellChecker::getSupportedServiceNames_Static());
		// acquire, because we return an interface pointer instead of a reference
		xFactory->acquire();
		pRet = xFactory.get();
	}
	return pRet;
}


///////////////////////////////////////////////////////////////////////////
