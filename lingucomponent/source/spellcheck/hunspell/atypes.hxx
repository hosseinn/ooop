#ifndef _ATYPES_HXX_
#define _ATYPES_HXX_

#ifndef HUNSPELL_WARNING
#ifdef HUNSPELL_WARNING_ON
#define HUNSPELL_WARNING fprintf
#else
// empty inline function to switch off warnings (instead of the C99 standard variadic macros)
static inline void HUNSPELL_WARNING(FILE *, const char *, ...) {}
#endif
#endif
 
// HUNSTEM def.
#define HUNSTEM

#include "csutil.hxx"
#include "hashmgr.hxx"

#define SETSIZE         256
#define CONTSIZE        65536
#define MAXWORDLEN      100
#define MAXWORDUTF8LEN  256

// affentry options
#define aeXPRODUCT      (1 << 0)
#define aeUTF8          (1 << 1)
#define aeALIASF        (1 << 2)
#define aeALIASM        (1 << 3)
#define aeINFIX         (1 << 4)

// compound options
#define IN_CPD_NOT   0
#define IN_CPD_BEGIN 1
#define IN_CPD_END   2
#define IN_CPD_OTHER 3

#define MAXLNLEN        8192

#define MINCPDLEN       3
#define MAXCOMPOUND     10

#define MAXACC          1000

#define FLAG unsigned short
#define FLAG_NULL 0x00
#define FREE_FLAG(a) a = 0

#define TESTAFF( a, b , c ) flag_bsearch((unsigned short *) a, (unsigned short) b, c)

struct affentry
{
   char * strip;
   char * appnd;
   unsigned char stripl;
   unsigned char appndl;
   char  numconds;
   char  opts;
   unsigned short aflag;
   union {
        char   base[SETSIZE];
        struct {
                char ascii[SETSIZE/2];
                char neg[8];
                char all[8];
                w_char * wchars[8];
                int wlen[8];
        } utf8;
   } conds;
#ifdef HUNSPELL_EXPERIMENTAL
   char *       morphcode;
#endif
   unsigned short * contclass;
   short        contclasslen;
};

struct mapentry {
  char * set;
  w_char * set_utf16;
  int len;
};

struct flagentry {
  FLAG * def;
  int len;
};

struct guessword {
  char * word;
  bool allow;
  char * orig;
};

#endif
