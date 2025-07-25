# Style

Alignment directly maps the alignment settings of the cells.

```python
class Alignment:
    horizontal: str = ""
    indent: int = 0
    justify_last_line: bool = False
    reading_order: int = 0
    relative_indent: int = 0
    shrink_to_fit: bool = False
    text_rotation: int = 0
    vertical: str = ""
    wrap_text: bool = False
```

Border directly maps the border settings of the cells.

```python
class Border:
    type: str = ""
    color: str = ""
    style: int = 0
```

Font directly maps the font settings of the fonts.

```python
class Font:
    bold: bool = False
    italic: bool = False
    underline: str = ""
    family: str = ""
    size: float = 0
    strike: bool = False
    color: str = ""
    color_indexed: int = 0
    color_theme: Optional[int] = None
    color_tint: float = 0
    vert_align: str = ""
    charset: Optional[int] = None
```

Fill directly maps the fill settings of the cells.

```python
class Fill:
    type: str = ""
    pattern: int = 0
    color: Optional[List[str]] = None
    shading: int = 0
```

Protection directly maps the protection settings of the cells.

```python
class Protection:
    hidden: bool = False
    locked: bool = False
```

Style directly maps the style settings of the cells.

```python
class Style:
    border: Optional[List[Border]] = None
    fill: Fill = Fill
    font: Optional[Font] = None
    alignment: Optional[Alignment] = None
    protection: Optional[Protection] = None
    num_fmt: int = 0
    decimal_places: Optional[int] = None
    custom_num_fmt: Optional[str] = None
    neg_red: bool = False
```

## Create style {#NewStyle}

```python
def new_style(style: Style) -> int
```

Create the style for cells by a given style options, and returns style index. The same style index can not be used across different workbook. This function is concurrency safe. Note that the `font.color` field uses an RGB color represented in `RRGGBB` hexadecimal notation.

### Border {#border}

The following table shows the border types used in `border.type` supported by excelize:

Type|Description|Type|Description
---|---|---|---
left|Left border|top|Top border
right|Right border|bottom|Bottom border
diagonalDown|Diagonal down border|diagonalUp|Diagonal up border

The following table shows the border styles used in `border.style` supported by excelize index number:

Index|Style|Line|Preview
---|---|---|---
0|None|0|
1|Continuous|1|!["Continuous"](https://xuri.me/excelize/images/style/border_01.png)
2|Continuous|2|!["Continuous"](https://xuri.me/excelize/images/style/border_02.png)
3|Dash|1|!["Dash"](https://xuri.me/excelize/images/style/border_03.png)
4|Dot|1|!["Dot"](https://xuri.me/excelize/images/style/border_04.png)
5|Continuous|3|!["Continuous"](https://xuri.me/excelize/images/style/border_05.png)
6|Double|3|!["Double"](https://xuri.me/excelize/images/style/border_06.png)
7|Continuous|0|!["Continuous"](https://xuri.me/excelize/images/style/border_07.png)
8|Dash|2|!["Dash"](https://xuri.me/excelize/images/style/border_08.png)
9|Dash Dot|1|!["Dash Dot"](https://xuri.me/excelize/images/style/border_09.png)
10|Dash Dot|2|!["Dash Dot"](https://xuri.me/excelize/images/style/border_10.png)
11|Dash Dot Dot|1|!["Dash Dot Dot"](https://xuri.me/excelize/images/style/border_11.png)
12|Dash Dot Dot|2|!["Dash Dot Dot"](https://xuri.me/excelize/images/style/border_12.png)
13|SlantDash Dot|2|!["SlantDash Dot"](https://xuri.me/excelize/images/style/border_13.png)

The following table shows the border styles used in `border.style` in the order shown in the Excel dialog:

Index|Preview|Index|Preview
---|---|---|---
0||12|!["Border 12"](https://xuri.me/excelize/images/style/border_12.png)
7|!["Border 7"](https://xuri.me/excelize/images/style/border_07.png)|13|!["Border 13"](https://xuri.me/excelize/images/style/border_13.png)
4|!["Border 4"](https://xuri.me/excelize/images/style/border_04.png)|10|!["Border 19"](https://xuri.me/excelize/images/style/border_10.png)
11|!["Border 11"](https://xuri.me/excelize/images/style/border_11.png)|8|!["Border 8"](https://xuri.me/excelize/images/style/border_08.png)
9|!["Border 9"](https://xuri.me/excelize/images/style/border_09.png)|2|!["Border 2"](https://xuri.me/excelize/images/style/border_02.png)
3|!["Border 3"](https://xuri.me/excelize/images/style/border_03.png)|5|!["Border 5"](https://xuri.me/excelize/images/style/border_05.png)
1|!["Border 1"](https://xuri.me/excelize/images/style/border_01.png)|6|!["Border 6"](https://xuri.me/excelize/images/style/border_06.png)

### Color fill {#shading}

The following table shows the shading styles used in `fill.shading` supported by excelize index number:

Index|Style|Index|Style
---|---|---|---
0|Horizontal|3|Diagonal down
1|Vertical|4|From corner
2|Diagonal Up|5|From center

### Pattern fill {#pattern}

The following table shows the pattern styles used in `fill.pattern` supported by excelize index number:

Index|Style|Index|Style
---|---|---|---
0|None|10|!["Pattern fill 10"](https://xuri.me/excelize/images/style/pattern_10.png)
1|!["Pattern fill 1"](https://xuri.me/excelize/images/style/pattern_01.png)|11|!["Pattern fill 11"](https://xuri.me/excelize/images/style/pattern_11.png)
2|!["Pattern fill 2"](https://xuri.me/excelize/images/style/pattern_02.png)|12|!["Pattern fill 12"](https://xuri.me/excelize/images/style/pattern_12.png)
3|!["Pattern fill 3"](https://xuri.me/excelize/images/style/pattern_03.png)|13|!["Pattern fill 13"](https://xuri.me/excelize/images/style/pattern_13.png)
4|!["Pattern fill 4"](https://xuri.me/excelize/images/style/pattern_04.png)|14|!["Pattern fill 14"](https://xuri.me/excelize/images/style/pattern_14.png)
5|!["Pattern fill 5"](https://xuri.me/excelize/images/style/pattern_05.png)|15|!["Pattern fill 15"](https://xuri.me/excelize/images/style/pattern_15.png)
6|!["Pattern fill 6"](https://xuri.me/excelize/images/style/pattern_06.png)|16|!["Pattern fill 16"](https://xuri.me/excelize/images/style/pattern_16.png)
7|!["Pattern fill 7"](https://xuri.me/excelize/images/style/pattern_07.png)|17|!["Pattern fill 17"](https://xuri.me/excelize/images/style/pattern_17.png)
8|!["Pattern fill 8"](https://xuri.me/excelize/images/style/pattern_08.png)|18|!["Pattern fill 18"](https://xuri.me/excelize/images/style/pattern_18.png)
9|!["Pattern fill 9](https://xuri.me/excelize/images/style/pattern_09.png)||

### Align {#align}

#### Indent

The `indent` is an integer value, where an increment of 1 represents 3 spaces. Indicates the number of spaces (of the normal style font) of indentation for text in a cell. The number of spaces to indent is calculated as following:

Number of spaces to indent = indent value * 3

For example, an indent value of 1 means that the text begins 3 space widths (of the normal style font) from the edge of the cell. Note: The width of one space character is defined by the font. Only left, right, and distributed horizontal alignments are supported.

#### Horizontal alignment

The following table shows the type of cells' horizontal alignment used in `alignment.horizontal`:

Type|Style
---|---
left|Left (indented)
center|Centered
right|Right (indented)
fill|Filling
justify|Justified
centerContinuous|Cross-column centered
distributed|Decentralized alignment (indented)

#### Vertical alignment

The following table shows the type of cells' vertical alignment used in `alignment.vertical`:

Type|Style
---|---
top|Top alignment
center|Centered
justify|Justified
distributed|Decentralized alignment

#### Reading order

`reading_order` is an uint64 value indicating whether the reading order of the cell is left-to-right, right-to-left, or context dependent. the valid value of this field was:

Value|Description
---|---
0 | Context Dependent - reading order is determined by scanning the text for the first non-whitespace character: if it is a strong right-to-left character, the reading order is right-to-left; otherwise, the reading order left-to-right.
1 | Left-to-Right: reading order is left-to-right in the cell, as in English.
2 | Right-to-Left: reading order is right-to-left in the cell, as in Hebrew.

#### Relative indent

`RelativeIndent` is an integer value to indicate the additional number of spaces of indentation to adjust for text in a cell.

### Font underline {#underline}

The following table shows the type of font underline style used in `font.underline`:

Type|Style
---|---
single|Single line
double|Double line

### Number format {#number_format}

Excel's built-in all languages formats (the `style.num_fmt` field) are shown in the following table:

Index|Type
---|---
0|`General`
1|`0`
2|`0.00`
3|`#,##0`
4|`#,##0.00`
5|`($#,##0_);($#,##0)`
6|`($#,##0_);[Red]($#,##0)`
7|`($#,##0.00_);($#,##0.00)`
8|`($#,##0.00_);[Red]($#,##0.00)`
9|`0%`
10|`0.00%`
11|`0.00E+00`
12|`# ?/?`
13|`# ??/??`
14|`mm-dd-yy`
15|`d-mmm-yy`
16|`d-mmm`
17|`mmm-yy`
18|`h:mm AM/PM`
19|`h:mm:ss AM/PM`
20|`h:mm`
21|`h:mm:ss`
22|`m/d/yy h:mm`
...|`...`
37|`(#,##0_);(#,##0)`
38|`(#,##0_);[Red](#,##0)`
39|`(#,##0.00_);(#,##0.00)`
40|`(#,##0.00_);[Red](#,##0.00)`
41|`_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`
42|`_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`
43|`_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`
44|`_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`
45|`mm:ss`
46|`[h]:mm:ss`
47|`mm:ss.0`
48|`##0.0E+0`
49|`@`

#### Traditional Chinese number format

Number format code in `zh-tw` language:

Index|Type
---|---
27|`[$-404]e/m/d`
28|`[$-404]e"年"m"月"d"日"`
29|`[$-404]e"年"m"月"d"日"`
30|`m/d/yy`
31|`yyyy"年"m"月"d"日"`
32|`hh"時"mm"分"`
33|`hh"時"mm"分"ss"秒"`
34|`上午/下午 hh"時"mm"分"`
35|`上午/下午 hh"時"mm"分"ss"秒"`
36|`[$-404]e/m/d`
50|`[$-404]e/m/d`
51|`[$-404]e"年"m"月"d"日"`
52|`上午/下午 hh"時"mm"分"`
53|`上午/下午 hh"時"mm"分"ss"秒"`
54|`[$-404]e"年"m"月"d"日"`
55|`上午/下午 hh"時"mm"分"`
56|`上午/下午 hh"時"mm"分"ss"秒"`
57|`[$-404]e/m/d`
58|`[$-404]e"年"m"月"d"日"`

#### Simplified Chinese number format

Number format code in `zh-cn` language:

Index|Type
---|---
27|`yyyy"年"m"月"`
28|`m"月"d"日"`
29|`m"月"d"日"`
30|`m-d-yy`
31|`yyyy"年"m"月"d"日"`
32|`h"时"mm"分"`
33|`h"时"mm"分"ss"秒"`
34|`上午/下午 h"时"mm"分"`
35|`上午/下午 h"时"mm"分"ss"秒`
36|`yyyy"年"m"月`
50|`yyyy"年"m"月`
51|`m"月"d"日`
52|`yyyy"年"m"月`
53|`m"月"d"日`
54|`m"月"d"日`
55|`上午/下午 h"时"mm"分`
56|`上午/下午 h"时"mm"分"ss"秒`
57|`yyyy"年"m"月`
58|`m"月"d"日"`

#### Japanese number format

Number format code in `ja-jp` language:

Index|Type
---|---
27|`[$-411]ge.m.d`
28|`[$-411]ggge"年"m"月"d"日`
29|`[$-411]ggge"年"m"月"d"日`
30|`m/d/y`
31|`yyyy"年"m"月"d"日`
32|`h"時"mm"分`
33|`h"時"mm"分"ss"秒`
34|`yyyy"年"m"月`
35|`m"月"d"日`
36|`[$-411]ge.m.d`
50|`[$-411]ge.m.d`
51|`[$-411]ggge"年"m"月"d"日`
52|`yyyy"年"m"月`
53|`m"月"d"日`
54|`[$-411]ggge"年"m"月"d"日`
55|`yyyy"年"m"月`
56|`m"月"d"日`
57|`[$-411]ge.m.d`
58|`[$-411]ggge"年"m"月"d"日"`

#### Korean number format

Number format code in `ko-kr` language:

Index|Type
---|---
27|`yyyy"年" mm"月" dd"日`
28|`mm-d`
29|`mm-d`
30|`mm-dd-y`
31|`yyyy"년" mm"월" dd"일`
32|`h"시" mm"분`
33|`h"시" mm"분" ss"초`
34|`yyyy-mm-d`
35|`yyyy-mm-d`
36|`yyyy"年" mm"月" dd"日`
50|`yyyy"年" mm"月" dd"日`
51|`mm-d`
52|`yyyy-mm-d`
53|`yyyy-mm-d`
54|`mm-d`
55|`yyyy-mm-d`
56|`yyyy-mm-d`
57|`yyyy"年" mm"月" dd"日`
58|`mm-dd`

#### Thai language number format

Number format code in `th-th` language:

Index|Type
---|---
59|`t`
60|`t0.0`
61|`t#,##`
62|`t#,##0.0`
67|`t0`
68|`t0.00`
69|`t# ?/`
70|`t# ??/?`
71|`ว/ด/ปปป`
72|`ว-ดดด-ป`
73|`ว-ดด`
74|`ดดด-ป`
75|`ช:น`
76|`ช:นน:ท`
77|`ว/ด/ปปปป ช:น`
78|`นน:ท`
79|`[ช]:นน:ท`
80|`นน:ทท.`
81|`d/m/bb`

### Currency format

Excelize built-in currency formats are shown in the following table, only support these types in the following table (Index number is used only for markup and is not used inside an Excel file and you can't get formatted value by the function [`get_cell_value`](cell.md#GetCellValue)) currently:

Index|Currency format
---|---
164|¥
165|$ English (United States)
166|$ Cherokee (United States)
167|$ Chinese (Singapore)
168|$ Chinese (Taiwan)
169|$ English (Australia)
170|$ English (Belize)
171|$ English (Canada)
172|$ English (Jamaica)
173|$ English (New Zealand)
174|$ English (Singapore)
175|$ English (Trinidad & Tobago)
176|$ English (U.S. Virgin Islands)
177|$ English (United States)
178|$ French (Canada)
179|$ Hawaiian (United States)
180|$ Malay (Brunei)
181|$ Quechua (Ecuador)
182|$ Spanish (Chile)
183|$ Spanish (Colombia)
184|$ Spanish (Ecuador)
185|$ Spanish (El Salvador)
186|$ Spanish (Mexico)
187|$ Spanish (Puerto Rico)
188|$ Spanish (United States)
189|$ Spanish (Uruguay)
190|£ English (United Kingdom)
191|£ Scottish Gaelic (United Kingdom)
192|£ Welsh (United Kindom)
193|¥ Chinese (China)
194|¥ Japanese (Japan)
195|¥ Sichuan Yi (China)
196|¥ Tibetan (China)
197|¥ Uyghur (China)
198|֏ Armenian (Armenia)
199|؋ Pashto (Afghanistan)
200|؋ Persian (Afghanistan)
201|৳ Bengali (Bangladesh)
202|៛ Khmer (Cambodia)
203|₡ Spanish (Costa Rica)
204|₦ Hausa (Nigeria)
205|₦ Igbo (Nigeria)
206|₩ Korean (South Korea)
207|₪ Hebrew (Israel)
208|₫ Vietnamese (Vietnam)
209|€ Basque (Spain)
210|€ Breton (France)
211|€ Catalan (Spain)
212|€ Corsican (France)
213|€ Dutch (Belgium)
214|€ Dutch (Netherlands)
215|€ English (Ireland)
216|€ Estonian (Estonia)
217|€ Euro (€ 123)
218|€ Euro (123 €)
219|€ Finnish (Finland)
220|€ French (Belgium)
221|€ French (France)
222|€ French (Luxembourg)
223|€ French (Monaco)
224|€ French (Réunion)
225|€ Galician (Spain)
226|€ German (Austria)
227|€ German (German)
228|€ German (Luxembourg)
229|€ Greek (Greece)
230|€ Inari Sami (Finland)
231|€ Irish (Ireland)
232|€ Italian (Italy)
233|€ Latin (Italy)
234|€ Latin, Serbian (Montenegro)
235|€ Larvian (Latvia)
236|€ Lithuanian (Lithuania)
237|€ Lower Sorbian (Germany)
238|€ Luxembourgish (Luxembourg)
239|€ Maltese (Malta)
240|€ Northern Sami (Finland)
241|€ Occitan (France)
242|€ Portuguese (Portugal)
243|€ Serbian (Montenegro)
244|€ Skolt Sami (Finland)
245|€ Slovak (Slovakia)
246|€ Slovenian (Slovenia)
247|€ Spanish (Spain)
248|€ Swedish (Finland)
249|€ Swiss German (France)
250|€ Upper Sorbian (Germany)
251|€ Western Frisian (Netherlands)
252|₭ Lao (Laos)
253|₮ Mongolian (Mongolia)
254|₮ Mongolian, Mongolian (Mongolia)
255|₱ English (Philippines)
256|₱ Filipino (Philippines)
257|₴ Ukrainian (Ukraine)
258|₸ Kazakh (Kazakhstan)
259|₹ Arabic, Kashmiri (India)
260|₹ English (India)
261|₹ Gujarati (India)
262|₹ Hindi (India)
263|₹ Kannada (India)
264|₹ Kashmiri (India)
265|₹ Konkani (India)
266|₹ Manipuri (India)
267|₹ Marathi (India)
268|₹ Nepali (India)
269|₹ Oriya (India)
270|₹ Punjabi (India)
271|₹ Sanskrit (India)
272|₹ Sindhi (India)
273|₹ Tamil (India)
274|₹ Urdu (India)
275|₺ Turkish (Turkey)
276|₼ Azerbaijani (Azerbaijan)
277|₼ Cyrillic, Azerbaijani (Azerbaijan)
278|₽ Russian (Russia)
279|₽ Sakha (Russia)
280|₾ Georgian (Georgia)
281|B/. Spanish (Panama)
282|Br Oromo (Ethiopia)
283|Br Somali (Ethiopia)
284|Br Tigrinya (Ethiopia)
285|Bs Quechua (Bolivia)
286|Bs Spanish (Bolivia)
287|BS. Spanish (Venezuela)
288|BWP Tswana (Botswana)
289|C$ Spanish (Nicaragua)
290|CA$ Latin, Inuktitut (Canada)
291|CA$ Mohawk (Canada)
292|CA$ Unified Canadian Aboriginal Syllabics, Inuktitut (Canada)
293|CFA French (Mali)
294|CFA French (Senegal)
295|CFA Fulah (Senegal)
296|CFA Wolof (Senegal)
297|CHF French (Switzerland)
298|CHF German (Liechtenstein)
299|CHF German (Switzerland)
300|CHF Italian (Switzerland)
301|CHF Romansh (Switzerland)
302|CLP Mapuche (Chile)
303|CN¥ Mongolian, Mongolian (China)
304|DZD Central Atlas Tamazight (Algeria)
305|FCFA French (Cameroon)
306|Ft Hungarian (Hungary)
307|G French (Haiti)
308|Gs. Spanish (Paraguay)
309|GTQ K'iche' (Guatemala)
310|HK$ Chinese (Hong Kong (China))
311|HK$ English (Hong Kong (China))
312|HRK Croatian (Croatia)
313|IDR English (Indonesia)
314|IQD Arbic, Central Kurdish (Iraq)
315|ISK Icelandic (Iceland)
316|K Burmese (Myanmar (Burma))
317|Kč Czech (Czech Republic)
318|KM Bosnian (Bosnia & Herzegovina)
319|KM Croatian (Bosnia & Herzegovina)
320|KM Latin, Serbian (Bosnia & Herzegovina)
321|kr Faroese (Faroe Islands)
322|kr Northern Sami (Norway)
323|kr Northern Sami (Sweden)
324|kr Norwegian Bokmål (Norway)
325|kr Norwegian Nynorsk (Norway)
326|kr Swedish (Sweden)
327|kr. Danish (Denmark)
328|kr. Kalaallisut (Greenland)
329|Ksh Swahili (kenya)
330|L Romanian (Moldova)
331|L Russian (Moldova)
332|L Spanish (Honduras)
333|Lekë Albanian (Albania)
334|MAD Arabic, Central Atlas Tamazight (Morocco)
335|MAD French (Morocco)
336|MAD Tifinagh, Central Atlas Tamazight (Morocco)
337|MOP$ Chinese (Macau (China))
338|MVR Divehi (Maldives)
339|Nfk Tigrinya (Eritrea)
340|NGN Bini (Nigeria)
341|NGN Fulah (Nigeria)
342|NGN Ibibio (Nigeria)
343|NGN Kanuri (Nigeria)
344|NOK Lule Sami (Norway)
345|NOK Southern Sami (Norway)
346|NZ$ Maori (New Zealand)
347|PKR Sindhi (Pakistan)
348|PYG Guarani (Paraguay)
349|Q Spanish (Guatemala)
350|R Afrikaans (South Africa)
351|R English (South Africa)
352|R Zulu (South Africa)
353|R$ Portuguese (Brazil)
354|RD$ Spanish (Dominican Republic)
355|RF Kinyarwanda (Rwanda)
356|RM English (Malaysia)
357|RM Malay (Malaysia)
358|RON Romanian (Romania)
359|Rp Indonesoan (Indonesia)
360|Rs Urdu (Pakistan)
361|Rs. Tamil (Sri Lanka)
362|RSD Latin, Serbian (Serbia)
363|RSD Serbian (Serbia)
364|RUB Bashkir (Russia)
365|RUB Tatar (Russia)
366|S/. Quechua (Peru)
367|S/. Spanish (Peru)
368|SEK Lule Sami (Sweden)
369|SEK Southern Sami (Sweden)
370|soʻm Latin, Uzbek (Uzbekistan)
371|soʻm Uzbek (Uzbekistan)
372|SYP Syriac (Syria)
373|THB Thai (Thailand)
374|TMT Turkmen (Turkmenistan)
375|US$ English (Zimbabwe)
376|ZAR Northern Sotho (South Africa)
377|ZAR Southern Sotho (South Africa)
378|ZAR Tsonga (South Africa)
379|ZAR Tswana (south Africa)
380|ZAR Venda (South Africa)
381|ZAR Xhosa (South Africa)
382|zł Polish (Poland)
383|ден Macedonian (Macedonia)
384|KM Cyrillic, Bosnian (Bosnia & Herzegovina)
385|KM Serbian (Bosnia & Herzegovina)
386|лв. Bulgarian (Bulgaria)
387|p. Belarusian (Belarus)
388|сом Kyrgyz (Kyrgyzstan)
389|сом Tajik (Tajikistan)
390|ج.م. Arabic (Egypt)
391|د.أ. Arabic (Jordan)
392|د.أ. Arabic (United Arab Emirates)
393|د.ب. Arabic (Bahrain)
394|د.ت. Arabic (Tunisia)
395|د.ج. Arabic (Algeria)
396|د.ع. Arabic (Iraq)
397|د.ك. Arabic (Kuwait)
398|د.ل. Arabic (Libya)
399|د.م. Arabic (Morocco)
400|ر Punjabi (Pakistan)
401|ر.س. Arabic (Saudi Arabia)
402|ر.ع. Arabic (Oman)
403|ر.ق. Arabic (Qatar)
404|ر.ي. Arabic (Yemen)
405|ریال Persian (Iran)
406|ل.س. Arabic (Syria)
407|ل.ل. Arabic (Lebanon)
408|ብር Amharic (Ethiopia)
409|रू Nepaol (Nepal)
410|රු. Sinhala (Sri Lanka)
411|ADP
412|AED
413|AFA
414|AFN
415|ALL
416|AMD
417|ANG
418|AOA
419|ARS
420|ATS
421|AUD
422|AWG
423|AZM
424|AZN
425|BAM
426|BBD
427|BDT
428|BEF
429|BGL
430|BGN
431|BHD
432|BIF
433|BMD
434|BND
435|BOB
436|BOV
437|BRL
438|BSD
439|BTN
440|BWP
441|BYR
442|BZD
443|CAD
444|CDF
445|CHE
446|CHF
447|CHW
448|CLF
449|CLP
450|CNY
451|COP
452|COU
453|CRC
454|CSD
455|CUC
456|CVE
457|CYP
458|CZK
459|DEM
460|DJF
461|DKK
462|DOP
463|DZD
464|ECS
465|ECV
466|EEK
467|EGP
468|ERN
469|ESP
470|ETB
471|EUR
472|FIM
473|FJD
474|FKP
475|FRF
476|GBP
477|GEL
478|GHC
479|GHS
480|GIP
481|GMD
482|GNF
483|GRD
484|GTQ
485|GYD
486|HKD
487|HNL
488|HRK
489|HTG
490|HUF
491|IDR
492|IEP
493|ILS
494|INR
495|IQD
496|IRR
497|ISK
498|ITL
499|JMD
500|JOD
501|JPY
502|KAF
503|KES
504|KGS
505|KHR
506|KMF
507|KPW
508|KRW
509|KWD
510|KYD
511|KZT
512|LAK
513|LBP
514|LKR
515|LRD
516|LSL
517|LTL
518|LUF
519|LVL
520|LYD
521|MAD
522|MDL
523|MGA
524|MGF
525|MKD
526|MMK
527|MNT
528|MOP
529|MRO
530|MTL
531|MUR
532|MVR
533|MWK
534|MXN
535|MXV
536|MYR
537|MZM
538|MZN
539|NAD
540|NGN
541|NIO
542|NLG
543|NOK
544|NPR
545|NTD
546|NZD
547|OMR
548|PAB
549|PEN
550|PGK
551|PHP
552|PKR
553|PLN
554|PTE
555|PYG
556|QAR
557|ROL
558|RON
559|RSD
560|RUB
561|RUR
562|RWF
563|SAR
564|SBD
565|SCR
566|SDD
567|SDG
568|SDP
569|SEK
570|SGD
571|SHP
572|SIT
573|SKK
574|SLL
575|SOS
576|SPL
577|SRD
578|SRG
579|STD
580|SVC
581|SYP
582|SZL
583|THB
584|TJR
585|TJS
586|TMM
587|TMT
588|TND
589|TOP
590|TRL
591|TRY
592|TTD
593|TWD
594|TZS
595|UAH
596|UGX
597|USD
598|USN
599|USS
600|UYI
601|UYU
602|UZS
603|VEB
604|VEF
605|VND
606|VUV
607|WST
608|XAF
609|XAG
610|XAU
611|XB5
612|XBA
613|XBB
614|XBC
615|XBD
616|XCD
617|XDR
618|XFO
619|XFU
620|XOF
621|XPD
622|XPF
623|XPT
624|XTS
625|XXX
626|YER
627|YUM
628|ZAR
629|ZMK
630|ZMW
631|ZWD
632|ZWL
633|ZWN
634|ZWR

Excelize support set custom number format for cell. For example, set number as date type in Uruguay (Spanish) format for `Sheet1!A6`:

<p align="center"><img width="612" src="https://xuri.me/excelize/en/images/number_format_01.png" alt="Set number format use Excelize for Python"></p>

```python
import excelize

try:
    f = excelize.new_file()
except RuntimeError as err:
    print(err)
    exit()
try:
    f.set_cell_value("Sheet1", "A6", 42920.5)
    style = f.new_style(
        excelize.Style(
            custom_num_fmt='[$-380A]dddd\\,\\ dd" de "mmmm" de "yyyy;@',
        )
    )
    f.set_cell_style("Sheet1", "A6", "A6", style)
    f.save_as("Book1.xlsx")
except RuntimeError as err:
    print(err)
finally:
    err = f.close()
    if err:
        print(err)
```

Cell `Sheet1!A6` in the Excel Application: `martes, 04 de Julio de 2017`

## Get style {#GetStyle}

```python
def get_style(style_id: int) -> Optional[Style]
```

Get style definition by given style ID.

## Set column style {#SetColStyle}

```python
def set_col_style(sheet: str, columns: str, style_id: int) -> None
```

Set style of columns by given worksheet name, columns range and style ID. Note that this will overwrite the existing styles for the columns, it won't append or merge style with existing styles.

For example set style of column `H` on `Sheet1`:

```python
try:
    f.set_col_style("Sheet1", "H", style)
except RuntimeError as err:
    print(err)
```

Set style of columns `C:F` on `Sheet1`:

```python
try:
    f.set_col_style("Sheet1", "C:F", style)
except RuntimeError as err:
    print(err)
```

## Get column style {#GetColStyle}

```python
def get_col_style(sheet: str, col: str) -> int
```

Get column style ID by given worksheet name and column name.

## Set row style {#SetRowStyle}

```python
def set_row_style(sheet: str, start: int, end: int, style_id: int) -> None
```

Set the style of rows by given worksheet name, row range, and style ID. Note that this will overwrite the existing styles for the rows, it won't append or merge style with existing styles.

For example set style of row 1 on `Sheet1`:

```python
try:
    f.set_row_style("Sheet1", 1, 1, style_id)
except RuntimeError as err:
    print(err)
```

Set style of rows 1 to 10 on `Sheet1`:

```python
try:
    f.set_row_style("Sheet1", 1, 10, style_id)
except RuntimeError as err:
    print(err)
```

## Set default font {#SetDefaultFont}

```python
def set_default_font(font_name: str) -> None
```

SetDefaultFont changes the default font in the workbook. The spreadsheet generated by excelize default font is Calibri.

## Get default font {#GetDefaultFont}

```python
def get_default_font() -> str
```

GetDefaultFont provides the default font name currently set in the workbook. The spreadsheet generated by excelize default font is Calibri.
