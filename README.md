# Simple README.md file

## FOREPLAY

This project was created with LibreOffice [^1]. A spreadsheet file CalcTest2.ods [^2] was created.

Double click on this and, it should be launched. Before, go in a safe place in your directory structure and ;-), proceed the next if not done before:

```
git clone https://github.com/cbushdor/myMakro.git
```

Double click on CalcTest2.ods but before go in the directory that contains this file and CalcTest2.ods.

Set your spreadsheet to launch macro if not already set! Et voìlà!


# Introduction
Fill these columns with date, labels and, their associated values (Asset or Passive).

Press the button and, calculus will be made for you.

# How to fill column(s)
* Date : Type date [^3] (format not specified yet)

* Label : String 

* Asset ( A ): Have a value

* Passive ( P ): Have a value

* Value: Can be a number (Real/Float) or _ [^4]

* Now value formats are checked (for A, P). An error is raised if format is not ok (check the upper line, or examples just below)! :-)

Watchout to the GOLDEN RULE:

> A vs P  if Asset has a value Passive MUST NOT (_ or 0)! It is one or the other (check the upper line, or examples just below).
> P vs A  if Passive has a value Asset MUST NOT (_ or 0)! It is one or the other (check the upper line, or examples just below).
> These rules are not implemented yet.

| Date | Label | Asset | Passive |
| ----------- | ----------- | ----------- | ----------- |
| 20230101 | Label 1 | 10 | _ |
| 20230101 | Label 2 | _  | 2|
| 20230101 | Label 3 | 9.1  | _ |


<!--
![This is the alt tag](./Screenshot_2023-01-02_at_03.07.02.png)
![This is the alt tag](./Screenshot_2023-01-02_at_03.07.38.png)
-->


# Improve macro
Tools > Organize Macros > Basic > Edit Macros

Menu Macro From > CalcTest2.ods > Standard > CalcSheet > Calc

Select Module and choose Calc()

```
REM Sum within one sheet
Sub CalcSheet(mySheet 	as String)
	MaSomme mySheet,"Asset","Asset sum" 			REM Calculus Column named Asset
	MaSomme mySheet,"Passive","Passive sum"			REM Calculus Column named Passive
	MaSomme mySheet,"My sum","Sums"					REM Calculus Column named Asset - Passive
End Sub

REM Start summup with one sheet
Sub Calc()
	CalcSheet("Sheet1")
End Sub
```
[![Everything Is AWESOME](https://i.imgur.com/Y3DVIlZ.png)](https://youtu.be/BcDxbIWh6qo)

[^1]: Version 7.4.2.3 check for latest news [^6]
[^2]: [Help](https://help.libreoffice.org/7.4/en-US/text/sbasic/shared/vbasupport.html?&DbPAR=SHARED&System=MAC) for writing scripts and, a [sheet cheat](https://documentation.libreoffice.org/assets/Uploads/Documentation/en/MACROS/RefCards/LibOBasic-3-Calc-Flat-A4-EN-v111.pdf)
[^3]: Format not specified yet
[^4]: _ [^5] is like 0 value
[^5]: _ must be alone
[^6]: Latest [news](https://wiki.documentfoundation.org/Main_Page)!
