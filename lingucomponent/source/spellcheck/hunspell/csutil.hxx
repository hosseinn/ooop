#ifndef __CSUTILHXX__
#define __CSUTILHXX__

// First some base level utility routines

#define NOCAP   0
#define INITCAP 1
#define ALLCAP  2
#define HUHCAP  3
#define HUHINITCAP  4

#define FIELD_STEM  "st:"
#define FIELD_POS   "po:"
#define FIELD_SUFF  "su:"
#define FIELD_PREF  "pr:"
#define FIELD_FREQ  "fr:"
#define FIELD_PHON  "ph:"
#define FIELD_HYPH  "hy:"
#define FIELD_COMP  "co:"

// default flags
#define ONLYUPCASEFLAG 65535

typedef struct {
    unsigned char l;
    unsigned char h;
} w_char;

#define w_char_eq(a,b) (((a).l == (b).l) && ((a).h == (b).h))

// convert UTF-16 characters to UTF-8
char * u16_u8(char * dest, int size, const w_char * src, int srclen);

// convert UTF-8 characters to UTF-16
int u8_u16(w_char * dest, int size, const char * src);

// sort 2-byte vector
void flag_qsort(unsigned short flags[], int begin, int end);

// binary search in 2-byte vector
int flag_bsearch(unsigned short flags[], unsigned short flag, int right);

// remove end of line char(s)
void   mychomp(char * s);

// duplicate string
char * mystrdup(const char * s);

// duplicate reverse of string
char * myrevstrdup(const char * s);

// parse into tokens with char delimiter
char * mystrsep(char ** sptr, const char delim);
// parse into tokens with char delimiter
char * mystrsep2(char ** sptr, const char delim);

// parse into tokens with char delimiter
char * mystrrep(char *, const char *, const char *);

// append s to ends of every lines in text
void strlinecat(char * lines, const char * s);

// tokenize into lines with new line
   int line_tok(const char * text, char *** lines);

// tokenize into lines with new line and uniq in place
   char * line_uniq(char * text);

// change \n to c in place
   char * line_join(char * text, char c);

// leave only last {[^}]*} pattern in string
   char * delete_zeros(char * morphout);

// reverse word
   int reverseword(char *);

// reverse word
   int reverseword_utf(char *);

// character encoding information
struct cs_info {
  unsigned char ccase;
  unsigned char clower;
  unsigned char cupper;
};

// two character arrays
struct replentry {
  char * pattern;
  char * pattern2;
};

// Unicode character encoding information
struct unicode_info {
  unsigned short c;
  unsigned short cupper;
  unsigned short clower;
};

struct unicode_info2 {
  char cletter;
  unsigned short cupper;
  unsigned short clower;
};

int initialize_utf_tbl();
void free_utf_tbl();
unsigned short unicodetoupper(unsigned short c, int langnum);
unsigned short unicodetolower(unsigned short c, int langnum);
int unicodeisalpha(unsigned short c);

struct enc_entry {
  const char * enc_name;
  struct cs_info * cs_table;
};

// language to encoding default map

struct lang_map {
  const char * lang;
  const char * def_enc;
  int num;
};

struct cs_info * get_current_cs(const char * es);

const char * get_default_enc(const char * lang);

// get language identifiers of language codes
int get_lang_num(const char * lang);

// get characters of the given 8bit encoding with lower- and uppercase forms
char * get_casechars(const char * enc);

// convert null terminated string to all caps using encoding
void enmkallcap(char * d, const char * p, const char * encoding);

// convert null terminated string to all little using encoding
void enmkallsmall(char * d, const char * p, const char * encoding);

// convert null terminated string to have intial capital using encoding
void enmkinitcap(char * d, const char * p, const char * encoding);

// convert null terminated string to all caps
void mkallcap(char * p, const struct cs_info * csconv);

// convert null terminated string to all little
void mkallsmall(char * p, const struct cs_info * csconv);

// convert null terminated string to have intial capital
void mkinitcap(char * p, const struct cs_info * csconv);

// convert first nc characters of UTF-8 string to little
void mkallsmall_utf(w_char * u, int nc, int langnum);

// convert first nc characters of UTF-8 string to capital
void mkallcap_utf(w_char * u, int nc, int langnum);

// get type of capitalization
int get_captype(char * q, int nl, cs_info *);

// get type of capitalization (UTF-8)
int get_captype_utf8(w_char * q, int nl, int langnum);

// strip all ignored characters in the string
void remove_ignored_chars_utf(char * word, unsigned short ignored_chars[], int ignored_len);

// strip all ignored characters in the string
void remove_ignored_chars(char * word, char * ignored_chars);

int parse_string(char * line, char ** out, const char * name);

int parse_array(char * line, char ** out,
        unsigned short ** out_utf16, int * out_utf16_len, const char * name, int utf8);

#endif
