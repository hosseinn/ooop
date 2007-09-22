/* Libhnj is dual licensed under LGPL and MPL. Boilerplate for both
 * licenses follows.
 */

/* LibHnj - a library for high quality hyphenation and justification
 * Copyright (C) 1998 Raph Levien, 
 * 	     (C) 2001 ALTLinux, Moscow (http://www.alt-linux.org), 
 *           (C) 2001 Peter Novodvorsky (nidd@cs.msu.su)
 *           (C) 2006 Németh László (nemeth at OOo)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Library General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Library General Public License for more details.
 *
 * You should have received a copy of the GNU Library General Public
 * License along with this library; if not, write to the 
 * Free Software Foundation, Inc., 59 Temple Place - Suite 330, 
 * Boston, MA  02111-1307  USA.
*/

/*
 * The contents of this file are subject to the Mozilla Public License
 * Version 1.0 (the "MPL"); you may not use this file except in
 * compliance with the MPL.  You may obtain a copy of the MPL at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the MPL is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the MPL
 * for the specific language governing rights and limitations under the
 * MPL.
 *
 */
#include <stdlib.h> /* for NULL, malloc */
#include <stdio.h>  /* for fprintf */
#include <string.h> /* for strdup */

#ifdef UNX
#include <unistd.h> /* for exit */
#endif

#include "sal/types.h"

#define noVERBOSE

#include "hnjalloc.h"
#include "hyphen.h"

static char *
hnj_strdup (const char *s)
{
  char *new;
  int l;

  l = strlen (s);
  new = hnj_malloc (l + 1);
  memcpy (new, s, l);
  new[l] = 0;
  return new;
}

/* remove cross-platform text line end characters */
void hnj_strchomp(char * s)
{
  int k = strlen(s);
  if ((k > 0) && ((*(s+k-1)=='\r') || (*(s+k-1)=='\n'))) *(s+k-1) = '\0';
  if ((k > 1) && (*(s+k-2) == '\r')) *(s+k-2) = '\0';
}

/* a little bit of a hash table implementation. This simply maps strings
   to state numbers */

typedef struct _HashTab HashTab;
typedef struct _HashEntry HashEntry;

/* A cheap, but effective, hack. */
#define HASH_SIZE 31627

struct _HashTab {
  HashEntry *entries[HASH_SIZE];
};

struct _HashEntry {
  HashEntry *next;
  char *key;
  int val;
};

/* a char* hash function from ASU - adapted from Gtk+ */
static unsigned int
hnj_string_hash (const char *s)
{
  const char *p;
  unsigned int h=0, g;
  for(p = s; *p != '\0'; p += 1) {
    h = ( h << 4 ) + *p;
    if ( 0 != ( g = h & 0xf0000000 ) ) {
      h = h ^ (g >> 24);
      h = h ^ g;
    }
  }
  return h /* % M */;
}

static HashTab *
hnj_hash_new (void)
{
  HashTab *hashtab;
  int i;

  hashtab = hnj_malloc (sizeof(HashTab));
  for (i = 0; i < HASH_SIZE; i++)
    hashtab->entries[i] = NULL;

  return hashtab;
}

static void
hnj_hash_free (HashTab *hashtab)
{
  int i;
  HashEntry *e, *next;

  for (i = 0; i < HASH_SIZE; i++)
    for (e = hashtab->entries[i]; e; e = next)
      {
	next = e->next;
	hnj_free (e->key);
	hnj_free (e);
      }

  hnj_free (hashtab);
}

/* assumes that key is not already present! */
static void
hnj_hash_insert (HashTab *hashtab, const char *key, int val)
{
  int i;
  HashEntry *e;

  i = hnj_string_hash (key) % HASH_SIZE;
  e = hnj_malloc (sizeof(HashEntry));
  e->next = hashtab->entries[i];
  e->key = hnj_strdup (key);
  e->val = val;
  hashtab->entries[i] = e;
}

/* return val if found, otherwise -1 */
static int
hnj_hash_lookup (HashTab *hashtab, const char *key)
{
  int i;
  HashEntry *e;
  i = hnj_string_hash (key) % HASH_SIZE;
  for (e = hashtab->entries[i]; e; e = e->next)
    if (!strcmp (key, e->key))
      return e->val;
  return -1;
}

/* Get the state number, allocating a new state if necessary. */
static int
hnj_get_state (HyphenDict *dict, HashTab *hashtab, const char *string)
{
  int state_num;

  state_num = hnj_hash_lookup (hashtab, string);

  if (state_num >= 0)
    return state_num;

  hnj_hash_insert (hashtab, string, dict->num_states);
  /* predicate is true if dict->num_states is a power of two */
  if (!(dict->num_states & (dict->num_states - 1)))
    {
      dict->states = hnj_realloc (dict->states,
				  (dict->num_states << 1) *
				  sizeof(HyphenState));
    }
  dict->states[dict->num_states].match = NULL;
  dict->states[dict->num_states].repl = NULL;
  dict->states[dict->num_states].fallback_state = -1;
  dict->states[dict->num_states].num_trans = 0;
  dict->states[dict->num_states].trans = NULL;
  return dict->num_states++;
}

/* add a transition from state1 to state2 through ch - assumes that the
   transition does not already exist */
static void
hnj_add_trans (HyphenDict *dict, int state1, int state2, char ch)
{
  int num_trans;

  num_trans = dict->states[state1].num_trans;
  if (num_trans == 0)
    {
      dict->states[state1].trans = hnj_malloc (sizeof(HyphenTrans));
    }
  else if (!(num_trans & (num_trans - 1)))
    {
      dict->states[state1].trans = hnj_realloc (dict->states[state1].trans,
						(num_trans << 1) *
						sizeof(HyphenTrans));
    }
  dict->states[state1].trans[num_trans].ch = ch;
  dict->states[state1].trans[num_trans].new_state = state2;
  dict->states[state1].num_trans++;
}

#ifdef VERBOSE
HashTab *global;

static char *
get_state_str (int state)
{
  int i;
  HashEntry *e;

  for (i = 0; i < HASH_SIZE; i++)
    for (e = global->entries[i]; e; e = e->next)
      if (e->val == state)
	return e->key;
  return NULL;
}
#endif

HyphenDict *
hnj_hyphen_load (const char *fn)
{
  HyphenDict *dict;
  HashTab *hashtab;
  FILE *f;
  char buf[MAX_CHARS];
  char word[MAX_CHARS];
  char pattern[MAX_CHARS];
  char * repl;
  signed char replindex;
  signed char replcut;
  int state_num = 0, last_state;
  int i, j;
  char ch;
  int found;
  HashEntry *e;

  f = fopen (fn, "r");
  if (f == NULL)
    return NULL;

  hashtab = hnj_hash_new ();
#ifdef VERBOSE
  global = hashtab;
#endif
  hnj_hash_insert (hashtab, "", 0);

  dict = hnj_malloc (sizeof(HyphenDict));
  dict->num_states = 1;
  dict->states = hnj_malloc (sizeof(HyphenState));
  dict->states[0].match = NULL;
  dict->states[0].repl = NULL;
  dict->states[0].fallback_state = -1;
  dict->states[0].num_trans = 0;
  dict->states[0].trans = NULL;

  /* read in character set info */
  for (i=0;i<MAX_NAME;i++) dict->cset[i]= 0;
  fgets(dict->cset,  sizeof(dict->cset),f);
  for (i=0;i<MAX_NAME;i++)
    if ((dict->cset[i] == '\r') || (dict->cset[i] == '\n'))
        dict->cset[i] = 0;
  dict->utf8 = (strcmp(dict->cset, "UTF-8") == 0);

  while (fgets (buf, sizeof(buf), f) != NULL)
    {
      if (buf[0] != '%')
	{
	  j = 0;
	  pattern[j] = '0';
          repl = strchr(buf, '/');
          replindex = 0;
          replcut = 0;
          if (repl) {
            char * index = strchr(repl + 1, ',');
            *repl = '\0';
            if (index) {
                char * index2 = strchr(index + 1, ',');
                *index = '\0';
				if (index2) {
                   *index2 = '\0';
                   replindex = (signed char) atoi(index + 1) - 1;
                   replcut = (signed char) atoi(index2 + 1);                
                }
            } else {
                hnj_strchomp(repl + 1);
                replindex = 0;
                replcut = (signed char) strlen(buf);
            }
            repl = hnj_strdup(repl + 1);
          }
	  for (i = 0; ((buf[i] > ' ') || (buf[i] < 0)); i++)
	    {
	      if (buf[i] >= '0' && buf[i] <= '9')
		pattern[j] = buf[i];
	      else
		{
		  word[j] = buf[i];
		  pattern[++j] = '0';
		}
	    }
	  word[j] = '\0';
	  pattern[j + 1] = '\0';

          i = 0;
	  if (!repl) {
	    /* Optimize away leading zeroes */
            for (; pattern[i] == '0'; i++);
          } else {
            if (*word == '.') i++;
            /* convert UTF-8 char. positions of discretionary hyph. replacements to 8-bit */
            if (dict->utf8) {
                int pu = -1;        /* unicode character position */
                int ps = -1;        /* unicode start position (original replindex) */
                int pc = (*word == '.') ? 1: 0; /* 8-bit character position */
                for (; pc < SAL_INT_CAST(int, strlen(word) + 1); pc++) {
                /* beginning of an UTF-8 character (not '10' start bits) */
                    if ((((unsigned char) word[pc]) >> 6) != 2) pu++;
                    if ((ps < 0) && (replindex == pu)) {
                        ps = replindex;
                        replindex = (signed char) pc;
                    }
                    if ((ps >= 0) && ((pu - ps) == replcut)) {
                        replcut = (signed char) (pc - replindex);
                        break;
                    }
                }
                if (*word == '.') replindex--;
            }
          }

#ifdef VERBOSE
	  printf ("word %s pattern %s, j = %d  repl: %s\n", word, pattern + i, j, repl);
#endif
	  found = hnj_hash_lookup (hashtab, word);
	  state_num = hnj_get_state (dict, hashtab, word);
	  dict->states[state_num].match = hnj_strdup (pattern + i);
	  dict->states[state_num].repl = repl;
	  dict->states[state_num].replindex = replindex;
          if (!replcut) {
            dict->states[state_num].replcut = (signed char) strlen(word);
          } else {
            dict->states[state_num].replcut = replcut;
          }

	  /* now, put in the prefix transitions */
          for (; found < 0 ;j--)
	    {
	      last_state = state_num;
	      ch = word[j - 1];
	      word[j - 1] = '\0';
	      found = hnj_hash_lookup (hashtab, word);
	      state_num = hnj_get_state (dict, hashtab, word);
	      hnj_add_trans (dict, state_num, last_state, ch);
	    }
	}
    }

  /* Could do unioning of matches here (instead of the preprocessor script).
     If we did, the pseudocode would look something like this:

     foreach state in the hash table
        foreach i = [1..length(state) - 1]
           state to check is substr (state, i)
           look it up
           if found, and if there is a match, union the match in.

     It's also possible to avoid the quadratic blowup by doing the
     search in order of increasing state string sizes - then you
     can break the loop after finding the first match.

     This step should be optional in any case - if there is a
     preprocessed rule table, it's always faster to use that.

*/

  /* put in the fallback states */
  for (i = 0; i < HASH_SIZE; i++)
    for (e = hashtab->entries[i]; e; e = e->next)
      {
	if (*(e->key)) for (j = 1; 1; j++)
	  {          
	    state_num = hnj_hash_lookup (hashtab, e->key + j);
	    if (state_num >= 0)
	      break;
	  }
        /* KBH: FIXME state 0 fallback_state should always be -1? */
	if (e->val)
	  dict->states[e->val].fallback_state = state_num;
      }
#ifdef VERBOSE
  for (i = 0; i < HASH_SIZE; i++)
    for (e = hashtab->entries[i]; e; e = e->next)
      {
	printf ("%d string %s state %d, fallback=%d\n", i, e->key, e->val,
		dict->states[e->val].fallback_state);
	for (j = 0; j < dict->states[e->val].num_trans; j++)
	  printf (" %c->%d\n", dict->states[e->val].trans[j].ch,
		  dict->states[e->val].trans[j].new_state);
      }
#endif

#ifndef VERBOSE
  hnj_hash_free (hashtab);
#endif
  
  fclose(f);
  return dict;
}

void hnj_hyphen_free (HyphenDict *dict)
{
  int state_num;
  HyphenState *hstate;

  for (state_num = 0; state_num < dict->num_states; state_num++)
    {
      hstate = &dict->states[state_num];
      if (hstate->match)
	hnj_free (hstate->match);
      if (hstate->repl)
	hnj_free (hstate->repl);
      if (hstate->trans)
	hnj_free (hstate->trans);
    }

  hnj_free (dict->states);

  hnj_free (dict);
}

#define MAX_WORD 256

int hnj_hyphen_hyphenate (HyphenDict *dict,
			   const char *word, int word_size,
			   char *hyphens)
{
  char prep_word_buf[MAX_WORD];
  char *prep_word;
  int i, j, k;
  int state;
  char ch;
  HyphenState *hstate;
  char *match;
  int offset;

  if (word_size + 3 < MAX_WORD)
    prep_word = prep_word_buf;
  else
    prep_word = hnj_malloc (word_size + 3);

  j = 0;
  prep_word[j++] = '.';
  
  for (i = 0; i < word_size; i++)
      prep_word[j++] = word[i];
      
  for (i = 0; i < j; i++)                                                       
    hyphens[i] = '0';    
  
  prep_word[j++] = '.';

  prep_word[j] = '\0';
#ifdef VERBOSE
  printf ("prep_word = %s\n", prep_word);
#endif

  /* now, run the finite state machine */
  state = 0;
  for (i = 0; i < j; i++)
    {
      ch = prep_word[i];
      for (;;)
	{

	  if (state == -1) {
            /* return 1; */
	    /*  KBH: FIXME shouldn't this be as follows? */
            state = 0;
            goto try_next_letter;
          }          

#ifdef VERBOSE
	  char *state_str;
	  state_str = get_state_str (state);

	  for (k = 0; k < i - strlen (state_str); k++)
	    putchar (' ');
	  printf ("%s", state_str);
#endif

	  hstate = &dict->states[state];
	  for (k = 0; k < hstate->num_trans; k++)
	    if (hstate->trans[k].ch == ch)
	      {
		state = hstate->trans[k].new_state;
		goto found_state;
	      }
	  state = hstate->fallback_state;
#ifdef VERBOSE
	  printf (" falling back, fallback_state %d\n", state);
#endif
	}
    found_state:
#ifdef VERBOSE
      printf ("found state %d\n",state);
#endif
      /* Additional optimization is possible here - especially,
	 elimination of trailing zeroes from the match. Leading zeroes
	 have already been optimized. */
      match = dict->states[state].match;
      /* replacing rules not handled by hyphen_hyphenate() */
      if (match && !dict->states[state].repl)
	{
	  offset = i + 1 - strlen (match);
#ifdef VERBOSE
	  for (k = 0; k < offset; k++)
	    putchar (' ');
	  printf ("%s\n", match);
#endif
	  /* This is a linear search because I tried a binary search and
	     found it to be just a teeny bit slower. */
	  for (k = 0; match[k]; k++)
	    if (hyphens[offset + k] < match[k])
	      hyphens[offset + k] = match[k];
	}

      /* KBH: we need this to make sure we keep looking in a word */
      /* for patterns even if the current character is not known in state 0 */
      /* since patterns for hyphenation may occur anywhere in the word */
      try_next_letter: ;

    }
#ifdef VERBOSE
  for (i = 0; i < j; i++)
    putchar (hyphens[i]);
  putchar ('\n');
#endif

  for (i = 0; i < j - 4; i++)
#if 0
    if (hyphens[i + 1] & 1)
      hyphens[i] = '-';
#else
    hyphens[i] = hyphens[i + 1];
#endif
  hyphens[0] = '0';
  for (; i < word_size; i++)
    hyphens[i] = '0';
  hyphens[word_size] = '\0';

  if (prep_word != prep_word_buf)
    hnj_free (prep_word);
  return 0;    
}

int hnj_hyphen_hyphenate2 (HyphenDict *dict,
			   const char *word, int word_size, char * hyphens,
			   char *hyphword, char *** rep, int ** pos, int ** cut)
{
  char prep_word_buf[MAX_WORD];
  char *prep_word;
  int i, j, k;
  int state;
  char ch;
  HyphenState *hstate;
  char *match;
  char *repl;
  signed char replindex;
  signed char replcut;
  int offset;
  int matchlen_buf[MAX_CHARS];
  int matchindex_buf[MAX_CHARS];
  char * matchrepl_buf[MAX_CHARS];
  int * matchlen;
  int * matchindex;
  char ** matchrepl;  
  int isrepl = 0;
  int nHyphCount;

  if (word_size + 3 < MAX_CHARS) {
    prep_word = prep_word_buf;
    matchlen = matchlen_buf;
    matchindex = matchindex_buf;
    matchrepl = matchrepl_buf;
  } else {
    prep_word = hnj_malloc (word_size + 3);
    matchlen = hnj_malloc ((word_size + 3) * sizeof(int));
    matchindex = hnj_malloc ((word_size + 3) * sizeof(int));
    matchrepl = hnj_malloc ((word_size + 3) * sizeof(char *));
  }

  j = 0;
  prep_word[j++] = '.';
  
  for (i = 0; i < word_size; i++)
      prep_word[j++] = word[i];

  prep_word[j++] = '.';
  prep_word[j] = '\0';

  for (i = 0; i < j; i++)
    hyphens[i] = '0';    

#ifdef VERBOSE
  printf ("prep_word = %s\n", prep_word);
#endif

  /* now, run the finite state machine */
  state = 0;
  for (i = 0; i < j; i++)
    {
      ch = prep_word[i];
      for (;;)
	{

	  if (state == -1) {
            /* return 1; */
	    /*  KBH: FIXME shouldn't this be as follows? */
            state = 0;
            goto try_next_letter;
          }          

#ifdef VERBOSE
	  char *state_str;
	  state_str = get_state_str (state);

	  for (k = 0; k < i - strlen (state_str); k++)
	    putchar (' ');
	  printf ("%s", state_str);
#endif

	  hstate = &dict->states[state];
	  for (k = 0; k < hstate->num_trans; k++)
	    if (hstate->trans[k].ch == ch)
	      {
		state = hstate->trans[k].new_state;
		goto found_state;
	      }
	  state = hstate->fallback_state;
#ifdef VERBOSE
	  printf (" falling back, fallback_state %d\n", state);
#endif
	}
    found_state:
#ifdef VERBOSE
      printf ("found state %d\n",state);
#endif
      /* Additional optimization is possible here - especially,
	 elimination of trailing zeroes from the match. Leading zeroes
	 have already been optimized. */
      match = dict->states[state].match;
      repl = dict->states[state].repl;
      replindex = dict->states[state].replindex;
      replcut = dict->states[state].replcut;
      /* replacing rules not handled by hyphen_hyphenate() */
      if (match)
	{
	  offset = i + 1 - strlen (match);
#ifdef VERBOSE
	  for (k = 0; k < offset; k++)
	    putchar (' ');
	  printf ("%s (%s)\n", match, repl);
#endif
          if (repl) {
            if (!isrepl) for(; isrepl < word_size; isrepl++) {
                matchrepl[isrepl] = NULL;
                matchindex[isrepl] = -1;
            }
            matchlen[offset + replindex] = replcut;
          }
	  /* This is a linear search because I tried a binary search and
	     found it to be just a teeny bit slower. */
	  for (k = 0; match[k]; k++) {
	    if ((hyphens[offset + k] < match[k])) {
	      hyphens[offset + k] = match[k];
              if (match[k]&1) {
                matchrepl[offset + k] = repl;
                if (repl && (k >= replindex) && (k <= replindex + replcut)) {
                    matchindex[offset + replindex] = offset + k;
                }
              }
            }
          }
          
	}

      /* KBH: we need this to make sure we keep looking in a word */
      /* for patterns even if the current character is not known in state 0 */
      /* since patterns for hyphenation may occur anywhere in the word */
      try_next_letter: ;

    }
#ifdef VERBOSE
  for (i = 0; i < j; i++)
    putchar (hyphens[i]);
  putchar ('\n');
#endif

  for (i = 0; i < j - 4; i++)
#if 0
    if (hyphens[i + 1] & 1)
      hyphens[i] = '-';
#else
    hyphens[i] = hyphens[i + 1];
#endif
  hyphens[0] = '0';
  for (; i < word_size; i++)
    hyphens[i] = '0';
  hyphens[word_size] = '\0';

  if (prep_word != prep_word_buf) {
    hnj_free (prep_word);
  }
        
       /* now create a new char string showing hyphenation positions */
       /* count the hyphens and allocate space for the new hypehanted string */
       nHyphCount = 0;
       for (i = 0; i < word_size; i++)
          if (hyphens[i]&1)
             nHyphCount++;
       j = 0;
       for (i = 0; i < word_size; i++) {
           if (isrepl && (matchindex[i] >= 0) && matchrepl[matchindex[i]]) {
                if (rep && pos && cut) {
                    if (!*rep && !*pos && !*cut) {
                        int k;
                        *rep = (char **) malloc(sizeof(char *) * word_size);
                        *pos = (int *) malloc(sizeof(int) * word_size);
                        *cut = (int *) malloc(sizeof(int) * word_size);
                        for (k = 0; k < word_size; k++) {
                            (*rep)[k] = NULL;
                            (*pos)[k] = 0;
                            (*cut)[k] = 0;
                        }
                    }
                    (*rep)[matchindex[i] - 1] = hnj_strdup(matchrepl[matchindex[i]]);
                    (*pos)[matchindex[i] - 1] = matchindex[i] - i;
                    (*cut)[matchindex[i] - 1] = matchlen[i];
                }
                if (hyphword) strcpy(hyphword + j, matchrepl[matchindex[i]]);
                j += strlen(matchrepl[matchindex[i]]);
                i += matchlen[i] - 1;
          } else if (hyphword) {
            hyphword[j++] = word[i];
            if ((hyphens[i]&1) && !(isrepl && ((i+1) < word_size) && 
                (matchindex[i+1] >= 0) && matchrepl[matchindex[i+1]])) hyphword[j++] = '=';
          }
       }
  if (hyphword) hyphword[j] = '\0';

  if (matchrepl != matchrepl_buf) {
    hnj_free (matchrepl);
    hnj_free (matchlen);
    hnj_free (matchindex);
  }

  if (!(dict->utf8)) return 0;

  if ((((unsigned char) word[0]) >> 6) == 2) {
    fprintf(stderr, "error - bad, non UTF-8 input: %s\n", word);
    return 1;
  }

  /* calculate UTF-8 character positions */
  j = -1;
  for (i = 0; i < word_size; i++) {
    /* beginning of an UTF-8 character (not '10' start bits) */
    if ((((unsigned char) word[i]) >> 6) != 2) j++;
    hyphens[j] = hyphens[i];
    if (rep && pos && cut && *rep && *pos && *cut) {
        int l = (*pos)[i];
        (*pos)[j] = 0;
        for (k = 0; k < l; k++) {
            if ((((unsigned char) word[i - k]) >> 6) != 2) (*pos)[j]++;
        }
        k = i - l + 1;
        l = k + (*cut)[i];
        (*cut)[j] = 0;        
        for (; k < l; k++) {
            if ((((unsigned char) word[k]) >> 6) != 2) (*cut)[j]++;
        }
        (*rep)[j] = (*rep)[i];
        if (j < i) {
            (*rep)[i] = NULL;
            (*pos)[i] = 0;
            (*cut)[i] = 0;
        }
    }
  }
  hyphens[j + 1] = '\0';
  return 0;
}

