# Simple readme file

This file is a LibreOffice [^1] file CalcTest2.ods [^2].

Double click on it and, it should be launched. 

Set your LibreOffice to launch macro.


# Introduction
Fill these columns with date, cost name and, its associated values (actif or Passif).

Press the button and, calculus will be made for you.

| Date | Label | Actif | Passif |
| ----------- | ----------- | ----------- | ----------- |
| 202301P1 | Label 1 | 10 | |
| 202301P1 | Label 2 |  | 2|

# How to feel column
* Date : Type date [^3] (format not specified yet)

* Label : String 

* Actif ( A ): Have a values

* Passif ( P ): Have a values

* Values: Can be a number (Real/Float) or _ [^4]  are both accepted

* Now values format are checked. An error is raised if format is not ok!

Watchout to the GOLDEN RULE:

> A vs P  if Actif as a value Passif MUST NOT (_ or 0) it sis one or the other.

> P vs A  if Passif as a value Actif MUST NOT (_ or 0) it sis one or the other.

> These rules are not implemented yet.

| Date | Label | Actif | Passif |
| ----------- | ----------- | ----------- | ----------- |
| 202301P1 | Label 1 | 10 | _ |
| 202301P1 | Label 2 | _  | 2|
| 202301P1 | Label 3 | 9.1  | _ |


<!--
![This is the alt tag](./Screenshot_2023-01-02_at_03.07.02.png)
![This is the alt tag](./Screenshot_2023-01-02_at_03.07.38.png)
-->


# Improve macro
Tools > Macro > Edit Macros

Go in Object catalog then choose CalcTest2.ods

Select Module and choose Calc()

```
Sub Calc()
	MaSomme("Sheet1","Actif","Reçu")
	MaSomme("Sheet1","Passif","Dépensé")
	MaSomme("Sheet1","Sommes","Total")
End Sub
```
[![Everything Is AWESOME](http://i.imgur.com/Ot5DWAW.png)](https://youtu.be/OXy5oTGC4Uk)

[^1]: Version 7.4.2.3 check for latest news [^6]
[^2]: [Manual](https://help.libreoffice.org/7.4/en-US/text/sbasic/shared/vbasupport.html?&DbPAR=SHARED&System=MAC) for writing VBA scripts
[^3]: Format not specified yet
[^4]: _ [^5] is like 0 value
[^5]: _ must be alone
[^6]: Latest [news](https://wiki.documentfoundation.org/Main_Page) for OO and VBA!
