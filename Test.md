
# Body Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents the body of a document or a section. Uma trying to clean up the table.

## Properties

| 

<span style="color: white;">Property</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 | 

<span style="color: white;">Req. Set</span>

 |
| 

**style**

 | 

string

 | 

Gets or sets the style used for the body. This is the name of the pre-installed or custom style.

 | 

1.1

 |
| 

**text**

 | 

string

 | 

Gets the text of the body. Use the insertText method to insert text. Read-only.

 | 

1.1

 |
| 

**type**

 | 

string

 | 

Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only. Possible values are: Unknown, MainDoc, Section, Header, Footer, TableCell.

 | 

1.3

 |

_See property access [examples.](#property-access-examples)_

## Relationships

| 

<span style="color: white;">Relationship</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 | 

<span style="color: white;">Req. Set</span>

 |
| 

**contentControls**

 | 

[ContentControlCollection](contentcontrolcollection.md)

 | 

Gets the collection of rich text content control objects in the body. Read-only.

 | 

1.1

 |
| 

**font**

 | 

[Font](font.md)

 | 

Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.

 | 

1.1

 |
| 

**inlinePictures**

 | 

[InlinePictureCollection](inlinepicturecollection.md)

 | 

Gets the collection of inlinePicture objects in the body. The collection does not include floating images. Read-only.

 | 

1.1

 |
| 

**lists**

 | 

[ListCollection](listcollection.md)

 | 

Gets the collection of list objects in the body. Read-only.

 | 

1.3

 |
| 

**paragraphs**

 | 

[ParagraphCollection](paragraphcollection.md)

 | 

Gets the collection of paragraph objects in the body. Read-only.

 | 

1.1

 |
| 

**parentBody**

 | 

[Body](body.md)

 | 

Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.

 | 

1.3

 |
| 

**parentContentControl**

 | 

[ContentControl](contentcontrol.md)

 | 

Gets the content control that contains the body. Returns null if there isn't a parent content control. Read-only.

 | 

1.1

 |
| 

**tables**

 | 

[TableCollection](tablecollection.md)

 | 

Gets the collection of table objects in the body. Read-only.

 | 

1.3

 |

## Methods

| 

<span style="color: white;">Method</span>

 | 

<span style="color: white;">Return Type</span>

 | 

<span style="color: white;">Description</span>

 | 

<span style="color: white;">Req. Set</span>

 |
| 

**[clear()](#clear)**

 | 

void

 | 

Clears the contents of the body object. The user can perform the undo operation on the cleared content.

 | 

1.1

 |
| 

**[getHtml()](#gethtml)**

 | 

string

 | 

Gets the HTML representation of the body object.

 | 

1.1

 |
| 

**[getOoxml()](#getooxml)**

 | 

string

 | 

Gets the OOXML (Office Open XML) representation of the body object.

 | 

1.1

 |
| 

**[getRange(rangeLocation: string)](#getrangerangelocation-string)**

 | 

[Range](range.md)

 | 

Gets the whole body, or the starting or ending point of the body, as a range.

 | 

1.3

 |
| 

**[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocat)**

 | 

void

 | 

Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.

 | 

1.1

 |
| 

**[insertContentControl()](#insertcontentcontrol)**

 | 

[ContentControl](contentcontrol.md)

 | 

Wraps the body object with a Rich Text content control.

 | 

1.1

 |
| 

**[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-i)**

 | 

[Range](range.md)

 | 

Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

 | 

1.1

 |
| 

**[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-st)**

 | 

[Range](range.md)

 | 

Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

 | 

1.1

 |
| 

**[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)](#insertinlinepicturefrombase64base64enco)**

 | 

[InlinePicture](inlinepicture.md)

 | 

Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.

 | 

1.2

 |
| 

**[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-)**

 | 

[Range](range.md)

 | 

Inserts OOXML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

 | 

1.1

 |
| 

**[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-ins)**

 | 

[Paragraph](paragraph.md)

 | 

Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.

 | 

1.1

 |
| 

**[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])](#inserttablerowcount-number-columncount-)**

 | 

[Table](table.md)

 | 

Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.

 | 

1.3

 |
| 

**[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-st)**

 | 

[Range](range.md)

 | 

Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

 | 

1.1

 |
| 

**[load(param: object)](#loadparam-object)**

 | 

void

 | 

Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

 | 

1.1

 |
| 

**[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-p)**

 | 

[SearchResultCollection](searchresultcollection.md)

 | 

Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.

 | 

1.1

 |
| 

**[select(selectionMode: string)](#selectselectionmode-string)**

 | 

void

 | 

Selects the body and navigates the Word UI to it.

 | 

1.1

 |

## Method Details

### clear()

Clears the contents of the body object. The user can perform the undo operation on the cleared content.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-c1">clear</span>();</pre>


#### Parameters

None

#### Returns

void

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to clear the contents of the body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-c1">clear</span>();</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Cleared the body contents.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Error:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Debug info:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **clear** method can be used to clear the contents of a document.

### getHtml()

Gets the HTML representation of the body object.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">getHtml</span>();</pre>


#### Parameters

None

#### Returns

string

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to get the HTML contents of the body.</span></pre>

<pre>    <span class="pl-k">var</span> bodyHTML <span class="pl-k">=</span> <span class="pl-smi">body</span>.<span class="pl-en">getHtml</span>();</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Body HTML contents:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-smi">bodyHTML</span>.<span class="pl-c1">value</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Error:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Debug info:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### getOoxml()

Gets the OOXML (Office Open XML) representation of the body object.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">getOoxml</span>();</pre>


#### Parameters

None

#### Returns

string

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to get the OOXML contents of the body.</span></pre>

<pre>    <span class="pl-k">var</span> bodyOOXML <span class="pl-k">=</span> <span class="pl-smi">body</span>.<span class="pl-en">getOoxml</span>();</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Body OOXML contents:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-smi">bodyOOXML</span>.<span class="pl-c1">value</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Error:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Debug info:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### getRange(rangeLocation: string)

Gets the whole body, or the starting or ending point of the body, as a range.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">getRange</span>(rangeLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**rangeLocation**

 | 

string

 | 

Optional. Optional. The range location can be 'Whole', 'Start' or 'End'. Possible values are: Whole, Start, End

 |

#### Returns

[Range](range.md)

### insertBreak(breakType: string, insertLocation: string)

Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertBreak</span>(breakType, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**breakType**

 | 

string

 | 

Required. The break type to add to the body. Possible values are: `<span style="font-size: 10pt;">Page</span>` Page break at the insertion point.,`<span style="font-size: 10pt;">Column</span>` Column break at the insertion point.,`<span style="font-size: 10pt;">Next</span>` Section break on next page.,`<span style="font-size: 10pt;">SectionContinuous</span>` New section without a corresponding page break.,`<span style="font-size: 10pt;">SectionEven</span>` Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.,`<span style="font-size: 10pt;">SectionOdd</span>` Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.,`<span style="font-size: 10pt;">Line</span>` Line break.,`<span style="font-size: 10pt;">LineClearLeft</span>` Line break.,`<span style="font-size: 10pt;">LineClearRight</span>` Line break.,`<span style="font-size: 10pt;">TextWrapping</span>` Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

void

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">ctx</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">ctx</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert a page break at the start of the document body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertBreak</span>(<span class="pl-smi">Word</span>.<span class="pl-smi">BreakType</span>.<span class="pl-smi">page</span>, <span class="pl-smi">Word</span>.<span class="pl-smi">InsertLocation</span>.<span class="pl-c1">start</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">ctx</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Added a page break at the start of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Error:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Debug info:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### insertContentControl()

Wraps the body object with a Rich Text content control.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertContentControl</span>();</pre>


#### Parameters

None

#### Returns

[ContentControl](contentcontrol.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to wrap the body in a content control.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertContentControl</span>();</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Wrapped the body in a content control.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### insertFileFromBase64(base64File: string, insertLocation: string)

Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertFileFromBase64</span>(base64File, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**base64File**

 | 

string

 | 

Required. The base64 encoded content of a .docx file.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[Range](range.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert base64 encoded .docx at the beginning of the content body.</span></pre>

<pre>    <span class="pl-c">// You will need to implement getBase64() to pass in a string of a base64 encoded docx file.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertFileFromBase64</span>(<span class="pl-en">getBase64</span>(), <span class="pl-smi">Word</span>.<span class="pl-smi">InsertLocation</span>.<span class="pl-c1">start</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Added base64 encoded text to the beginning of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### insertHtml(html: string, insertLocation: string)

Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertHtml</span>(html, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**html**

 | 

string

 | 

Required. The HTML to be inserted in the document.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[Range](range.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert HTML in to the beginning of the body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertHtml</span>(<span class="pl-pds">'</span><span class="pl-s"><strong>This is text inserted with body.insertHtml()</strong></span><span class="pl-pds">'</span>, <span class="pl-smi">Word</span>.<span class="pl-smi">InsertLocation</span>.<span class="pl-c1">start</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">HTML added to the beginning of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)

Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertInlinePictureFromBase64</span>(base64EncodedImage, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**base64EncodedImage**

 | 

string

 | 

Required. The base64 encoded image to be inserted in the body.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[InlinePicture](inlinepicture.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert OOXML in to the beginning of the body.</span></pre>

<pre>    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">OOXML added to the beginning of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


_Additional information_

Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML. The [Word-Add-in-DocumentAssembly][body.insertOoxml] sample shows how you can use this API to assemble a document.

### insertOoxml(ooxml: string, insertLocation: string)

Inserts OOXML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertOoxml</span>(ooxml, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**ooxml**

 | 

string

 | 

Required. The OOXML to be inserted.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: string)

Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertParagraph</span>(paragraphText, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**paragraphText**

 | 

string

 | 

Required. The paragraph text to be inserted.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[Paragraph](paragraph.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert the paragraph at the end of the document body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertParagraph</span>(<span class="pl-pds">'</span><span class="pl-s">Content of a new paragraph</span><span class="pl-pds">'</span>, <span class="pl-smi">Word</span>.<span class="pl-smi">InsertLocation</span>.<span class="pl-smi">end</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Paragraph added at the end of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


_Additional information_ The [Word-Add-in-DocumentAssembly][body.insertParagraph] sample shows how you can use the insertParagraph method to assemble a document.

### insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])

Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertTable</span>(rowCount, columnCount, insertLocation, values);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**rowCount**

 | 

number

 | 

Required. The number of rows in the table.

 |
| 

**columnCount**

 | 

number

 | 

Required. The number of columns in the table.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |
| 

**values**

 | 

string[][]

 | 

Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

 |

#### Returns

[Table](table.md)

### insertText(text: string, insertLocation: string)

Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-en">insertText</span>(text, insertLocation);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**text**

 | 

string

 | 

Required. Text to be inserted.

 |
| 

**insertLocation**

 | 

string

 | 

Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `<span style="font-size: 10pt;">Before</span>` Add content before the contents of the calling object.,`<span style="font-size: 10pt;">After</span>` Add content after the contents of the calling object.,`<span style="font-size: 10pt;">Start</span>` Prepend content to the contents of the calling object.,`<span style="font-size: 10pt;">End</span>` Append content to the contents of the calling object.,`<span style="font-size: 10pt;">Replace</span>` Replace the contents of the current object.

 |

#### Returns

[Range](range.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to insert text in to the beginning of the body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-en">insertText</span>(<span class="pl-pds">'</span><span class="pl-s">This is text inserted with body.insertText()</span><span class="pl-pds">'</span>, <span class="pl-smi">Word</span>.<span class="pl-smi">InsertLocation</span>.<span class="pl-c1">start</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Text added to the beginning of the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### load(param: object)

Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax

<div>

<pre><span class="pl-smi">object</span>.<span class="pl-c1">load</span>(param);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**param**

 | 

object

 | 

Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.

 |

#### Returns

void

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to load font and style information for the document body.</span></pre>

<pre>    <span class="pl-smi">context</span>.<span class="pl-c1">load</span>(body, <span class="pl-pds">'</span><span class="pl-s">font/size, font/name, font/color, style</span><span class="pl-pds">'</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-c">// Show the results of the load method. Here we show the</span></pre>

<pre>        <span class="pl-c">// property values on the body object.</span></pre>

<pre>        <span class="pl-k">var</span> results <span class="pl-k">=</span> <span class="pl-pds">'</span><span class="pl-s">Font size:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">size</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Font name:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">name</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Font color:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">color</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Body style:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-c1">style</span>;</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(results);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)

Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-c1">search</span>(searchText, searchOptions);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**searchText**

 | 

string

 | 

Required. The search text.

 |
| 

**searchOptions**

 | 

ParamTypeStrings.SearchOptions

 | 

Optional. Optional. Options for the search.

 |

#### Returns

[SearchResultCollection](searchresultcollection.md)

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to search the document.</span></pre>

<pre>    <span class="pl-k">var</span> searchResults <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>.<span class="pl-c1">search</span>(<span class="pl-pds">'</span><span class="pl-s">video</span><span class="pl-pds">'</span>, {matchCase<span class="pl-k">:</span> <span class="pl-c1">false</span>});</pre>

<pre>    <span class="pl-c">// Queue a commmand to load the results.</span></pre>

<pre>    <span class="pl-smi">context</span>.<span class="pl-c1">load</span>(searchResults, <span class="pl-pds">'</span><span class="pl-s">text, font</span><span class="pl-pds">'</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-k">var</span> results <span class="pl-k">=</span> <span class="pl-pds">'</span><span class="pl-s">Found count:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">searchResults</span>.<span class="pl-smi">items</span>.<span class="pl-c1">length</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; we highlighted the results.</span><span class="pl-pds">'</span>;</pre>

<pre>        <span class="pl-c">// Queue a command to change the font for each found item.</span></pre>

<pre>        <span class="pl-k">for</span> (<span class="pl-k">var</span> i <span class="pl-k">=</span> <span class="pl-c1">0</span>; i <span class="pl-k"><</span> <span class="pl-smi">searchResults</span>.<span class="pl-smi">items</span>.<span class="pl-c1">length</span>; i<span class="pl-k">++</span>) {</pre>

<pre>          <span class="pl-smi">searchResults</span>.<span class="pl-smi">items</span>[i].<span class="pl-smi">font</span>.<span class="pl-c1">color</span> <span class="pl-k">=</span> <span class="pl-pds">'</span><span class="pl-s">#FF0000</span><span class="pl-pds">'</span>    <span class="pl-c">// Change color to Red</span></pre>

<pre>          <span class="pl-smi">searchResults</span>.<span class="pl-smi">items</span>[i].<span class="pl-smi">font</span>.<span class="pl-smi">highlightColor</span> <span class="pl-k">=</span> <span class="pl-pds">'</span><span class="pl-s">#FFFF00</span><span class="pl-pds">'</span>;</pre>

<pre>          <span class="pl-smi">searchResults</span>.<span class="pl-smi">items</span>[i].<span class="pl-smi">font</span>.<span class="pl-smi">bold</span> <span class="pl-k">=</span> <span class="pl-c1">true</span>;</pre>

<pre>        }</pre>

<pre>        <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>        <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>        <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>            <span class="pl-en">console</span>.<span class="pl-c1">log</span>(results);</pre>

<pre>        });</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


_Additional information_ The [Word-Add-in-DocumentAssembly][body.search] sample provides another example of how to search a document.

### select(selectionMode: string)

Selects the body and navigates the Word UI to it.

#### Syntax

<div>

<pre><span class="pl-smi">bodyObject</span>.<span class="pl-c1">select</span>(selectionMode);</pre>


#### Parameters

| 

<span style="color: white;">Parameter</span>

 | 

<span style="color: white;">Type</span>

 | 

<span style="color: white;">Description</span>

 |
| 

**selectionMode**

 | 

string

 | 

Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default. Possible values are: Select, Start, End

 |

#### Returns

void

#### Examples

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to select the document body. The Word UI will</span></pre>

<pre>    <span class="pl-c">// move to the selected document body.</span></pre>

<pre>    <span class="pl-smi">body</span>.<span class="pl-c1">select</span>();</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Selected the document body.</span><span class="pl-pds">'</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


### Property access examples

_Get the text property on the body object_

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to load the text in document body.</span></pre>

<pre>    <span class="pl-smi">context</span>.<span class="pl-c1">load</span>(body, <span class="pl-pds">'</span><span class="pl-s">text</span><span class="pl-pds">'</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Body contents:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-c1">text</span>);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Error:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">"</span><span class="pl-s">Debug info:</span> <span class="pl-pds">"</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>

<pre>});</pre>


_Get the style and the font size, font name, and font color properties on the body object._

<div>

<pre><span class="pl-c">// Run a batch operation against the Word object model.</span></pre>

<pre><span class="pl-smi">Word</span>.<span class="pl-en">run</span>(<span class="pl-k">function</span> (<span class="pl-smi">context</span>) {</pre>

<pre>    <span class="pl-c">// Create a proxy object for the document body.</span></pre>

<pre>    <span class="pl-k">var</span> body <span class="pl-k">=</span> <span class="pl-smi">context</span>.<span class="pl-smi">document</span>.<span class="pl-c1">body</span>;</pre>

<pre>    <span class="pl-c">// Queue a commmand to load font and style information for the document body.</span></pre>

<pre>    <span class="pl-smi">context</span>.<span class="pl-c1">load</span>(body, <span class="pl-pds">'</span><span class="pl-s">font/size, font/name, font/color, style</span><span class="pl-pds">'</span>);</pre>

<pre>    <span class="pl-c">// Synchronize the document state by executing the queued commands,</span></pre>

<pre>    <span class="pl-c">// and return a promise to indicate task completion.</span></pre>

<pre>    <span class="pl-k">return</span> <span class="pl-smi">context</span>.<span class="pl-en">sync</span>().<span class="pl-en">then</span>(<span class="pl-k">function</span> () {</pre>

<pre>        <span class="pl-c">// Show the results of the load method. Here we show the</span></pre>

<pre>        <span class="pl-c">// property values on the body object.</span></pre>

<pre>        <span class="pl-k">var</span> results <span class="pl-k">=</span> <span class="pl-pds">'</span><span class="pl-s">Font size:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">size</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Font name:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">name</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Font color:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-smi">font</span>.<span class="pl-c1">color</span> <span class="pl-k">+</span></pre>

<pre>                      <span class="pl-pds">'</span><span class="pl-s">; Body style:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-smi">body</span>.<span class="pl-c1">style</span>;</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(results);</pre>

<pre>    });</pre>

<pre>})</pre>

<pre>.<span class="pl-en">catch</span>(<span class="pl-k">function</span> (<span class="pl-smi">error</span>) {</pre>

<pre>    <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Error:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(error));</pre>

<pre>    <span class="pl-k">if</span> (error <span class="pl-k">instanceof</span> <span class="pl-smi">OfficeExtension</span>.<span class="pl-smi">Error</span>) {</pre>

<pre>        <span class="pl-en">console</span>.<span class="pl-c1">log</span>(<span class="pl-pds">'</span><span class="pl-s">Debug info:</span> <span class="pl-pds">'</span> <span class="pl-k">+</span> <span class="pl-c1">JSON</span>.<span class="pl-en">stringify</span>(<span class="pl-smi">error</span>.<span class="pl-smi">debugInfo</span>));</pre>

<pre>    }</pre>


});
