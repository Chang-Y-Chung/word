{smcl}
{title:Title}

{phang}
{bf:word} {hline 2} Simply create a word document

{title:Syntax}

     Open a word document

{p 8 19 2}
{cmd:word} {cmd:open} [{cmd:using}] {it:filename}[{cmd:,} {cmd:replace}]

     Close the word document

{p 8 19 2}
{cmd:word} {cmd:close}

     Add an image to the word document

{p 8 19 2}
{cmd:word} {cmd:add} {cmd:image} [
{cmd:using}] {it:filename}
[, {cmd:link(}{it:number}{cmd:)}
{cmd:cx(}{it:width}{cmd:)}
{cmd:cy(}{it:height}{cmd:)}]

{title:Description}

{pstd}
{cmd:word} is a series of wrappers around some of the mata
{help mf__docx:[M-5] _docx*()} functions. {cmd:word} command
makes it simple to create a word document without switching
to mata. The syntax resembles that of {help [P] file}
command without the file handle part.

{title:Remarks}

{pstd}
{cmd:word} opens only one word document at a time, and supports
a subset of mata {help mf__docx:[M-5] _docx*()} functions.

{pstd}
{cmd:word add image} does not work on the following platforms:
Stata for Solaris (both SPARC and x86-64), Stata for Mac running
Mavericks, or console Stata for Mac.

{title:Examples}

{phang}{cmd:. word open test.docx, replace}{p_end}
{phang}{cmd:. word add image test.png}{p_end}
{phang}{cmd:. word close}{p_end}

