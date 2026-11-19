# Formula function coverage

> Written by `bun scripts/formula-coverage.ts` (do not edit by hand: run it again after adding a function). It compares the engine's function registry (`src/core/xlsx/formula/functions.ts`) with the function lists of [Excel](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical) (read 2026-11-19) and [Google Sheets](https://support.google.com/docs/table/25273?hl=en) (read 2026-11-19). `--refresh` downloads the two lists again.

The lists are taken as published: a function a help center has not listed is counted as not being there (Google Sheets computes `XMATCH`, its list does not have it), and neither list tells how much of a function the engine computes: only whether it has one of that name. What is known to differ is under [Limits](#limits).

## Summary

| Reference | Functions | Computed | Missing | Coverage |
| --- | ---: | ---: | ---: | ---: |
| Excel | 521 | 504 | 17 | 97% |
| Google Sheets | 515 | 504 | 11 | 98% |
| In both | 468 | 466 | 2 | 100% |
| In either | 568 | 542 | 26 | 95% |

The engine computes **542** of them (and `ANCHORARRAY` `SINGLE`, which neither list names); **26** are missing, all of them out of reach (see below). A formula that uses a missing one is not computed: it keeps the value the file holds, which no longer follows the cells it reads, and shows `#NAME?` when typed in the editor.

## By category

| Category | Functions | Computed | Missing | Coverage |
| --- | ---: | ---: | ---: | ---: |
| Math | 83 | 83 | 0 | 100% |
| Statistical | 151 | 151 | 0 | 100% |
| Text | 53 | 51 | 2 | 96% |
| Date | 26 | 26 | 0 | 100% |
| Logical | 19 | 19 | 0 | 100% |
| Lookup | 41 | 39 | 2 | 95% |
| Info | 24 | 22 | 2 | 92% |
| Array | 2 | 2 | 0 | 100% |
| Financial | 55 | 55 | 0 | 100% |
| Engineering | 57 | 57 | 0 | 100% |
| Database | 12 | 12 | 0 | 100% |
| Operator | 15 | 15 | 0 | 100% |
| Parser | 5 | 5 | 0 | 100% |
| Web | 9 | 3 | 6 | 33% |
| Google | 5 | 2 | 3 | 40% |
| AI | 1 | 0 | 1 | 0% |
| Cube | 7 | 0 | 7 | 0% |
| Add-in | 3 | 0 | 3 | 0% |

Categories are those of Excel, shortened; a function only Google Sheets has keeps the category it has there (*Array*, *Operator*, *Parser*, *Google*, *AI*). The old names Excel lists under *Compatibility* are counted with what they compute.

## What the engine evaluates besides functions

- **Names inside a formula**: `LET` and `LAMBDA` (with optional parameters, `ISOMITTED`, a LAMBDA called at once or passed to `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, a function named without being called: `BYROW(rows, SUM)`).
- **Dynamic arrays**: a formula that gives several values fills the cells next to it; `#SPILL!` when a cell is in the way; `A1#` for the cells a formula fills, `@` for the one value of a range that is the formula's own. Saved as Excel saves them (an array formula marked dynamic), so they stay dynamic there and here.
- **Defined names**, of the workbook and local to a sheet, **structured references** (`Table1[Qty]`, `Table1[@Qty]`, `[@Qty]`, `Table1[[#Headers],[Qty]]`, `Table1[#All]`) and **3-D references** (`First:Last!A1`).

## Out of reach

Not to be written: the value the file holds is the best there is.

| Function | Category | Excel | Google Sheets | Notes |
| --- | --- | --- | --- | --- |
| `DETECTLANGUAGE` | Text | [Microsoft 365](https://support.microsoft.com/en-us/excel/functions/detectlanguage-function) | [yes](https://support.google.com/docs/answer/3093278) |  |
| `TRANSLATE` | Text | [Microsoft 365](https://support.microsoft.com/en-us/excel/functions/translate-function) |  |  |
| `IMAGE` | Lookup | [2024](https://support.microsoft.com/en-us/excel/functions/image-function) | [yes](https://support.google.com/docs/answer/3093333) |  |
| `RTD` | Lookup | [yes](https://support.microsoft.com/en-us/excel/functions/rtd-function) |  |  |
| `INFO` | Info | [yes](https://support.microsoft.com/en-us/excel/functions/info-function) |  |  |
| `STOCKHISTORY` | Info | [yes](https://support.microsoft.com/en-us/excel/functions/stockhistory-function) |  |  |
| `IMPORTDATA` | Web |  | [yes](https://support.google.com/docs/answer/3093335) |  |
| `IMPORTFEED` | Web |  | [yes](https://support.google.com/docs/answer/3093337) |  |
| `IMPORTHTML` | Web |  | [yes](https://support.google.com/docs/answer/3093339) |  |
| `IMPORTRANGE` | Web |  | [yes](https://support.google.com/docs/answer/3093340) |  |
| `IMPORTXML` | Web |  | [yes](https://support.google.com/docs/answer/3093342) |  |
| `WEBSERVICE` | Web | [2013](https://support.microsoft.com/en-us/excel/functions/webservice-function) |  |  |
| `GOOGLEFINANCE` | Google |  | [yes](https://support.google.com/docs/answer/3093281) |  |
| `GOOGLETRANSLATE` | Google |  | [yes](https://support.google.com/docs/answer/3093331) |  |
| `SPARKLINE` | Google |  | [yes](https://support.google.com/docs/answer/3093289) |  |
| `AI` | AI |  | [yes](https://support.google.com/docs/answer/15877199) |  |
| `CUBEKPIMEMBER` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubekpimember-function) |  |  |
| `CUBEMEMBER` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubemember-function) |  |  |
| `CUBEMEMBERPROPERTY` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubememberproperty-function) |  |  |
| `CUBERANKEDMEMBER` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cuberankedmember-function) |  |  |
| `CUBESET` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubeset-function) |  |  |
| `CUBESETCOUNT` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubesetcount-function) |  |  |
| `CUBEVALUE` | Cube | [yes](https://support.microsoft.com/en-us/excel/functions/cubevalue-function) |  |  |
| `CALL` | Add-in | [yes](https://support.microsoft.com/en-us/excel/functions/call-function) |  |  |
| `EUROCONVERT` | Add-in | [yes](https://support.microsoft.com/en-us/excel/functions/euroconvert-function) |  |  |
| `REGISTER.ID` | Add-in | [yes](https://support.microsoft.com/en-us/excel/functions/register-id-function) |  |  |

## Limits

Functions the engine computes, and what it does not do as their application does:

| Functions | Limit |
| --- | --- |
| `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` | the model Excel names (exponential triple smoothing, additive), fitted by least squares: Excel does not publish how it fits its own, so a forecast may differ from Excel's in its last digits |
| `GROUPBY` `PIVOTBY` | one function at a time (not an array of functions side by side); PIVOTBY takes one column of values |
| `QUERY` | select, where, group by, order by, limit, offset, label; no `pivot`; `format` and `options` are read and left aside |
| `FILTERXML` | paths with `/`, `//`, `*`, `@attribute`, `text()` and predicates on position, attributes and children: not the whole of XPath |
| `REGEXTEST` `REGEXEXTRACT` `REGEXREPLACE` `REGEXMATCH` | the regular expressions of JavaScript, close to the PCRE2 of Excel and the RE2 of Google Sheets, not the same in every corner |
| `ASC` `DBCS` `JIS` `FINDB` `LEFTB` `LENB` `MIDB` `REPLACEB` `RIGHTB` `SEARCHB` | as Excel without a double-byte language set: the byte functions count characters |
| `PHONETIC` | the text itself: the phonetic reading a file keeps for Japanese text is not read |
| `TO_DATE` `TO_DOLLARS` `TO_PERCENT` `TO_PURE_NUMBER` | give the number as it is: how a cell shows it is the cell's format |
| `CELL` | `address`, `col`, `row`, `contents`, `type`, `format`, `color`, `parentheses`, `prefix`, `protect`, `width`; `filename` is empty |
| `GETPIVOTDATA` | reads the values a pivot table of the file shows (the engine does not compute pivot tables) |
| `SUBTOTAL` `AGGREGATE` | a row hidden inside a filter or a table counts as filtered out; hiding rows does not by itself compute the formulas again |
| `BESSELI` `BESSELJ` `BESSELK` `BESSELY` | computed to full precision: the examples of the help center show the fewer digits of older versions of Excel |

The engine opens the formats of Excel, and Google Sheets saves a function Excel does not have as `IFERROR(__xludf.DUMMYFUNCTION("…"), value)`: such a file opens with its values. The functions of Google Sheets count for formulas typed in the editor.

## Computed today

- **Math** (83): `ABS` `ACOS` `ACOSH` `ACOT` `ACOTH` `AGGREGATE` `ARABIC` `ASIN` `ASINH` `ATAN` `ATAN2` `ATANH` `BASE` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `COMBIN` `COMBINA` `COS` `COSH` `COT` `COTH` `COUNTUNIQUE` `CSC` `CSCH` `DECIMAL` `DEGREES` `EVEN` `EXP` `FACT` `FACTDOUBLE` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `GCD` `INT` `ISO.CEILING` `LCM` `LN` `LOG` `LOG10` `MDETERM` `MINVERSE` `MMULT` `MOD` `MROUND` `MULTINOMIAL` `MUNIT` `ODD` `PERCENTOF` `PI` `POWER` `PRODUCT` `QUOTIENT` `RADIANS` `RAND` `RANDARRAY` `RANDBETWEEN` `ROMAN` `ROUND` `ROUNDDOWN` `ROUNDUP` `SEC` `SECH` `SEQUENCE` `SERIESSUM` `SIGN` `SIN` `SINH` `SQRT` `SQRTPI` `SUBTOTAL` `SUM` `SUMIF` `SUMIFS` `SUMPRODUCT` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `TAN` `TANH` `TRUNC`

- **Statistical** (151): `AVEDEV` `AVERAGE` `AVERAGE.WEIGHTED` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `BETA.DIST` `BETA.INV` `BETADIST` `BETAINV` `BINOM.DIST` `BINOM.DIST.RANGE` `BINOM.INV` `BINOMDIST` `CHIDIST` `CHIINV` `CHISQ.DIST` `CHISQ.DIST.RT` `CHISQ.INV` `CHISQ.INV.RT` `CHISQ.TEST` `CHITEST` `CONFIDENCE` `CONFIDENCE.NORM` `CONFIDENCE.T` `CORREL` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `COVAR` `COVARIANCE.P` `COVARIANCE.S` `CRITBINOM` `DEVSQ` `EXPON.DIST` `EXPONDIST` `F.DIST` `F.DIST.RT` `F.INV` `F.INV.RT` `F.TEST` `FDIST` `FINV` `FISHER` `FISHERINV` `FORECAST` `FORECAST.ETS` `FORECAST.ETS.CONFINT` `FORECAST.ETS.SEASONALITY` `FORECAST.ETS.STAT` `FORECAST.LINEAR` `FREQUENCY` `FTEST` `GAMMA` `GAMMA.DIST` `GAMMA.INV` `GAMMADIST` `GAMMAINV` `GAMMALN` `GAMMALN.PRECISE` `GAUSS` `GEOMEAN` `GROWTH` `HARMEAN` `HYPGEOM.DIST` `HYPGEOMDIST` `INTERCEPT` `KURT` `LARGE` `LINEST` `LOGEST` `LOGINV` `LOGNORM.DIST` `LOGNORM.INV` `LOGNORMDIST` `MARGINOFERROR` `MAX` `MAXA` `MAXIFS` `MEDIAN` `MIN` `MINA` `MINIFS` `MODE` `MODE.MULT` `MODE.SNGL` `NEGBINOM.DIST` `NEGBINOMDIST` `NORM.DIST` `NORM.INV` `NORM.S.DIST` `NORM.S.INV` `NORMDIST` `NORMINV` `NORMSDIST` `NORMSINV` `PEARSON` `PERCENTILE` `PERCENTILE.EXC` `PERCENTILE.INC` `PERCENTRANK` `PERCENTRANK.EXC` `PERCENTRANK.INC` `PERMUT` `PERMUTATIONA` `PHI` `POISSON` `POISSON.DIST` `PROB` `QUARTILE` `QUARTILE.EXC` `QUARTILE.INC` `RANK` `RANK.AVG` `RANK.EQ` `RSQ` `SKEW` `SKEW.P` `SLOPE` `SMALL` `STANDARDIZE` `STDEV` `STDEV.P` `STDEV.S` `STDEVA` `STDEVP` `STDEVPA` `STEYX` `T.DIST` `T.DIST.2T` `T.DIST.RT` `T.INV` `T.INV.2T` `T.TEST` `TDIST` `TINV` `TREND` `TRIMMEAN` `TTEST` `VAR` `VAR.P` `VAR.S` `VARA` `VARP` `VARPA` `WEIBULL` `WEIBULL.DIST` `Z.TEST` `ZTEST`

- **Text** (51): `ARRAYTOTEXT` `ASC` `BAHTTEXT` `CHAR` `CLEAN` `CODE` `CONCAT` `CONCATENATE` `DBCS` `DOLLAR` `EXACT` `FIND` `FINDB` `FIXED` `JIS` `JOIN` `LEFT` `LEFTB` `LEN` `LENB` `LOWER` `MID` `MIDB` `NUMBERVALUE` `PHONETIC` `PROPER` `REGEXEXTRACT` `REGEXMATCH` `REGEXREPLACE` `REGEXTEST` `REPLACE` `REPLACEB` `REPT` `RIGHT` `RIGHTB` `SEARCH` `SEARCHB` `SPLIT` `SUBSTITUTE` `T` `TEXT` `TEXTAFTER` `TEXTBEFORE` `TEXTJOIN` `TEXTSPLIT` `TRIM` `UNICHAR` `UNICODE` `UPPER` `VALUE` `VALUETOTEXT`

- **Date** (26): `DATE` `DATEDIF` `DATEVALUE` `DAY` `DAYS` `DAYS360` `EDATE` `EOMONTH` `EPOCHTODATE` `HOUR` `ISOWEEKNUM` `MINUTE` `MONTH` `NETWORKDAYS` `NETWORKDAYS.INTL` `NOW` `SECOND` `TIME` `TIMEVALUE` `TODAY` `WEEKDAY` `WEEKNUM` `WORKDAY` `WORKDAY.INTL` `YEAR` `YEARFRAC`

- **Logical** (19): `AND` `BYCOL` `BYROW` `FALSE` `IF` `IFERROR` `IFNA` `IFS` `LAMBDA` `LET` `MAKEARRAY` `MAP` `NOT` `OR` `REDUCE` `SCAN` `SWITCH` `TRUE` `XOR`

- **Lookup** (39): `ADDRESS` `AREAS` `CHOOSE` `CHOOSECOLS` `CHOOSEROWS` `COLUMN` `COLUMNS` `DROP` `EXPAND` `FILTER` `FORMULATEXT` `GETPIVOTDATA` `GROUPBY` `HLOOKUP` `HSTACK` `HYPERLINK` `INDEX` `INDIRECT` `LOOKUP` `MATCH` `OFFSET` `PIVOTBY` `ROW` `ROWS` `SORT` `SORTBY` `SORTN` `TAKE` `TOCOL` `TOROW` `TRANSPOSE` `TRIMRANGE` `UNIQUE` `VLOOKUP` `VSTACK` `WRAPCOLS` `WRAPROWS` `XLOOKUP` `XMATCH`

- **Info** (22): `CELL` `ERROR.TYPE` `ISBLANK` `ISDATE` `ISEMAIL` `ISERR` `ISERROR` `ISEVEN` `ISFORMULA` `ISLOGICAL` `ISNA` `ISNONTEXT` `ISNUMBER` `ISODD` `ISOMITTED` `ISREF` `ISTEXT` `N` `NA` `SHEET` `SHEETS` `TYPE`

- **Array** (2): `ARRAY_CONSTRAIN` `FLATTEN`

- **Financial** (55): `ACCRINT` `ACCRINTM` `AMORDEGRC` `AMORLINC` `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNCD` `COUPNUM` `COUPPCD` `CUMIPMT` `CUMPRINC` `DB` `DDB` `DISC` `DOLLARDE` `DOLLARFR` `DURATION` `EFFECT` `FV` `FVSCHEDULE` `INTRATE` `IPMT` `IRR` `ISPMT` `MDURATION` `MIRR` `NOMINAL` `NPER` `NPV` `ODDFPRICE` `ODDFYIELD` `ODDLPRICE` `ODDLYIELD` `PDURATION` `PMT` `PPMT` `PRICE` `PRICEDISC` `PRICEMAT` `PV` `RATE` `RECEIVED` `RRI` `SLN` `SYD` `TBILLEQ` `TBILLPRICE` `TBILLYIELD` `VDB` `XIRR` `XNPV` `YIELD` `YIELDDISC` `YIELDMAT`

- **Engineering** (57): `BESSELI` `BESSELJ` `BESSELK` `BESSELY` `BIN2DEC` `BIN2HEX` `BIN2OCT` `BITAND` `BITLSHIFT` `BITOR` `BITRSHIFT` `BITXOR` `COMPLEX` `CONVERT` `DEC2BIN` `DEC2HEX` `DEC2OCT` `DELTA` `ERF` `ERF.PRECISE` `ERFC` `ERFC.PRECISE` `GESTEP` `HEX2BIN` `HEX2DEC` `HEX2OCT` `IMABS` `IMAGINARY` `IMARGUMENT` `IMCONJUGATE` `IMCOS` `IMCOSH` `IMCOT` `IMCOTH` `IMCSC` `IMCSCH` `IMDIV` `IMEXP` `IMLN` `IMLOG` `IMLOG10` `IMLOG2` `IMPOWER` `IMPRODUCT` `IMREAL` `IMSEC` `IMSECH` `IMSIN` `IMSINH` `IMSQRT` `IMSUB` `IMSUM` `IMTAN` `IMTANH` `OCT2BIN` `OCT2DEC` `OCT2HEX`

- **Database** (12): `DAVERAGE` `DCOUNT` `DCOUNTA` `DGET` `DMAX` `DMIN` `DPRODUCT` `DSTDEV` `DSTDEVP` `DSUM` `DVAR` `DVARP`

- **Operator** (15): `ADD` `DIVIDE` `EQ` `GT` `GTE` `ISBETWEEN` `LT` `LTE` `MINUS` `MULTIPLY` `NE` `POW` `UMINUS` `UNARY_PERCENT` `UPLUS`

- **Parser** (5): `TO_DATE` `TO_DOLLARS` `TO_PERCENT` `TO_PURE_NUMBER` `TO_TEXT`

- **Web** (3): `ENCODEURL` `FILTERXML` `ISURL`

- **Google** (2): `ARRAYFORMULA` `QUERY`

- **In neither list** (2): `ANCHORARRAY` (how a file writes the `#` operator: `A1#`), `SINGLE` (how a file writes the `@` operator)
