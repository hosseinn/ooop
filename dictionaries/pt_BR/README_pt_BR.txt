Autor/Author: Raimundo Moura <raimundomoura@openoffice.org>

pt-BR: Este dicion?rio est? em desenvolvimento por Raimundo Moura e sua equipe. Ele est? licenciado sob os termos da Licen?a P?blica Geral Menor vers?o 2.1 (LGPLv2.1), como publicado pela Free Software Foundation. Os cr?ditos est?o dispon?veis em http://www.broffice.org/creditos e voc? pode encontrar novas vers?es em http://www.broffice.org/verortografico.

en-US: This dictionary is under development by Raimundo Moura and his team. It is licensed under the terms of the GNU Lesser General Public License version 2.1 (LGPLv2.1), as published by the Free Software Foundation. The credits are available at http://www.broffice.org/creditos and you can find new releases at http://www.broffice.org/verortografico.


Copyright (C) 2006 - 2009 por/by Raimundo Santos Moura <raimundomoura@openoffice.org>

=============
APRESENTA??O
=============

O Projeto Verificador Ortogr?fico do BrOffice.org ? um projeto
colaborativo desenvolvido pela comunidade Brasileira.
A rela??o completa dos colaboradores deste projeto est? em:
http://www.broffice.org.br/creditos

***********************************************************************
* Este ? um dicion?rio para corre??o ortogr?fica da l?ngua Portuguesa *
* para o Myspell.                                                     *
* Este programa ? livre e pode ser redistribu?do e/ou modificado nos  *
* termos da GNU Lesser General Public License (LGPL) vers?o 2.1.      *
*                                                                     *
***********************************************************************

======================
SOBRE ESTA ATUALIZA??O
======================

. Altera??o da regra 'B' para inclus?o de plural para sufixos '?s'
  Exemplo: vi?s - vieses;
. Corre??o da regra 'a' para conjuga??o correta do verbo 'cuspir'.
  N?o estava contemplado 'cospem';
. Corre??o da regra 'm' para verbos terminados em quir/guir nas formas de
  ?nclises e mes?clises
. Inclus?o do comando BREAK para permitir a verifica??o correta dos compostos;
. Inclus?o do comando MAXNGRAMSUGS para limitar o n?mero de sugest?es;
. Atualiza??o da regra REP melhorando as op??es de sugest?o;
. Inclus?o de 'hidrossanit?rio', colabora??o de Gilmar Grespan;
. Inclus?o de 'tropeirismo', colabora??o de Tiago Hillebrandt;
. Corre??o da regra 'a' para verbos para inclus?o de superlativos;
. Inclus?o de: marquetagem, lasqueira, esquerdopata, latinoide, colabora??o Edson Costa;
. Corre??o da regra de sugest?es 'REP';
. Inclus?o de: sudoestino e Sanepar, colabora??o Marcos Vin?cius Piccinini   
. Corre??o da conjuga??o de verbos nas formas de ?nclises e mes?clises;
. Inclus?o de regras para composi??o de paises;
. Inclus?o das siglas dos partidos pol?ticos brasileiros;
. refor?o do prefixo ex-; 
. Inclus?o de: precursoramente, neoconstitucionalismo, alopoiese, alopoi?tico,
  autopoi?tico, BrOffice, paradoxiza??o, programaticidade, sistemismo, comteano,
  durkheimiano e luhmanniano. Colabora??o Pablo Feitosa;
. Exclus?o de 'fundamenais'. Colabora??o de Jo?o Paulo Vinha Bittar;
  Judicializa??o.  Colabora??o Pablo Feitosa;
. Inclus?o de: Judicializa??o, externaliza, externalizado, externaliza??o, externalizei, etc.
  Paradoxiza??o, desparadoxiza??o, procedimentalizar, procedimentaliza??o , 
  procedimentalizo, etc. Colabora??o Pablo Feitosa;
. Altera??es no Divsilab. Colabora??o Flavio Figueiredo Cardoso;
. Aplica??o das regras de pref?xo: sub, super, auto, re, inter, etc. na forma??o de novos
  compostos.  
. Inclus?o de:  racistoide, melequento, burraldo, entrega??o, prum, pruns
  (contra??o para+artigo), comezinha, yakuz? (m?fia japonesa) e figura?a. 
  Exclus?o de: 'p?ozinhos' no plural est? incorreto. Colabora??o de Edson Costa.
. Inclus?o de 'dimetilsulfato' e 'dietilsulfato'. Colabora??o Luis Alcides Brandini De Boni;


=======================================================
COMO INSTALAR O VERIFICADOR BRASILEIRO NO BROFFICE.ORG
=======================================================

Copie os arquivos pt_BR.dic e pt_BR.aff para o diret?rio <BrOffice.org>
/share/dict/ooo, onde <BrOffice.org> ? o diret?rio em que o programa 
foi instalado.

No Windows, normalmente, o caminho ? este: 
C:\Arquivos de programas\BrOffice.org 2.0\share\dict\ooo, e no  Linux
/opt/BrOffice.org/share/dict/ooo/.

No mesmo diret?rio, localize o arquivo dictionary.lst. Abra-o com um
editor de textos e acrescente a seguinte linha ao final(se n?o
existir):

DICT pt BR pt_BR

? necess?rio reiniciar o BrOffice, inclusive o in?cio r?pido da vers?o
para Windows que fica na barra de tarefas, para que o corretor
funcione.

===================
D?VIDAS FREQUENTES
===================

Os arquivos foram copiados mas o Verificador n?o est? funcionando.
O Verificador Ortogr?fico n?o deve estar configurado corretamente,
isto pode estar ocorrendo por um dos seguintes motivos:

1- O dicion?rio provavelmente n?o est? instalado.

Para se certificar de que est? utilizando o idioma correto confira como
est?o as informa??es em: Ferramentas >> Op??es >>   Configura??es de
Idioma >> Idiomas. O item Ocidental deve apresentar o dicion?rio
selecionado (deve aparecer um logo "Abc" do lado do idioma).

Se n?o estiver Portugu?s (Brasil) mude para esse idioma. Ap?s
configurado clique em 'OK'.
Feche o BrOffice, inclusive o Iniciador R?pido,  e em seguida reabra-o;


2 - O verificador n?o est? configurado para verificar texto ao digitar.
Neste caso confira como est?o as informa??es em: 

(At? a Vers?o 3.0.X)
Ferramentas >> Op??es>> Configura??es de Idiomas >> Recursos de Verifica??o
Ortogr?fica e, no campo op??es deste formul?rio marque a op??o 'Verificar 
texto ao digitar';

(Vers?o 3.1 em diante)
Ferramentas >> Op??es >> Configura??es de Idiomas >> Recursos para reda??o e,
no campo op??es deste formul?rio marque a op??o 'Verificar ortografia ao digitar


Novas atualiza??es estar?o dispon?veis no site do BrOffice.Org, na
p?gina do Verificador Ortogr?fico.

http://www.openoffice.org.br/?q=verortografico


============
INTRODUCTION
============

The BrOffice.org Orthography Checker is a colaborative project developed
by the Brazilian community.
The complete list of participants in this project is at
http://www.broffice.org.br/creditos

***********************************************************************
* This is a dictionary for orthography correction for the Portuguese  *
* language for Myspell.                                               *
* This is a free program and it can be redistributed and/or           *
* modified under the terms of the GNU Lesser General Public License   *
* (LGPL) version 2.1.                                                 *
*                                                                     *
***********************************************************************

=================
ABOUT THIS UPDATE
=================

==============================================================
HOW TO INSTALL THE BRAZILIAN ORTOGRAPH CHECKER IN BROFFICE.ORG
==============================================================

Copy the files pt_BR.dic and pt_BR.aff to the directory <BrOffice.org>
/share/dict/ooo, where <BrOffice.org> is the directory where the software
has been installed.

In Windows, usually, the path is
C:\Arquivos de programas\BrOffice.org 2.0\share\dict\ooo, and in GNU/Linux
/opt/BrOffice.org/share/dict/ooo/.

In the same directory, locate the file dictionary.lst. Open it with a
text editor e add the following line to the end of the file (if it is
not already there):

DICT pt BR pt_BR

It is necessary to restart BrOffice, including the fast start for the Windows version
that resides on the task bar, in order to have the orthography checker to work.


==========================
FREQUENTLY ASKED QUESTIONS
==========================

The files have been copied but the checker is not working. The orthography checker may not be
configured correctly, this may be due to one of the following reasons:

1- The dictionary is probably not installed.

To make sure that you are using the right language, check the information at
Ferramentas >> Op??es >>  Configura??es de Idioma >> Idiomas.
The item "Ocidental" must present the selected dictionary (a logo "Abc" should
appear beside the language).
If the language selected is not "Portugu?s (Brasil)" change to this language.
After the configuration is correct, click on 'OK'.
Close BrOffice and the fast start, and open it afterwards;

2 - The checker is not configured to verify the orthography on typing. For this

problem, check the information at
"Ferramentas >> Op??es >> Configura??es de Idiomas >> Recursos de Verifica??o Ortogr?fica"
and, in the field "Op??es" of this form, check the option ''Verificar texto ao digitar';

New updates will be available at the BrOffice.Org website, on the page of the
Orthography Checker.

http://www.broffice.org/verortografico

