*! version 0.0.1 cyc 2013-10-25

global word_handle ""
global word_filename ""

prog def word
  version 13.1
  gettoken cmd 0: 0
  if ("`cmd'" == "") exit 198 // invalid syntax
  else word_`cmd' `0'
end

prog def word_add
  gettoken cmd 0: 0
  if ("`cmd'" == "") exit 198 // invalid syntax
  else word_add_`cmd' `0'
end

prog def word_add_image
  // this function is not supported on following:
  //   Stata for Solaris (both SPARC and x86-64)
  //   Stata for Mac running on Mac OS X 10.9 (Mavericks)
  //   console Stata for Mac
  // in any of the above, the function returns error code -16525

  has_handle
  if !r(has_handle) error 603 // file could not be opened

  cap syntax using/ [, link(real 0.0) cx(real 0.0) cy(real 0.0)]
  if _rc cap syntax anything(name=using1) [, link(real 0.0) cx(real 0.0) cy(real 0.0)]
  if _rc error 198

  is_blank `"`using'"'
  local using = cond(r(is_blank), `"`using1'"', `"`using'"')

  mata_func `"_docx_image_add($word_handle, `"`using'"', `link', `cx', `cy' )"'
  word_save
end

prog def word_add_text
  has_handle
  if !r(has_handle) error 603 // file could not be opened

  syntax [anything(name=str)] [, Style(string)]

  unquote `"`str'"'
  local str `"`r(unquoted)'"'
  unquote `"`style'"'
  local style `"`r(unquoted)'"'

  mata_func `"_docx_paragraph_new_styledtext($word_handle, `"`str'"', `"`style'"')"'
  word_save
end

prog def word_close
  has_handle
  if !r(has_handle) error 198 // syntax error

  word_save
  mata_func `"_docx_close($word_handle)"'
  macro drop word_*
end

prog def word_open
  has_handle
  if r(has_handle) error 681 // too many open files. a la -16521

  cap syntax using/ [, replace]
  if _rc cap syntax anything(name=using1) [, replace]
  if _rc exit 198 // invalid syntax

  is_blank `"`using'"'
  local using = cond(r(is_blank), `"`using1'"', `"`using'"')

  scalar has_type = regexm(`"`using'"', "(\.[^.]+)$")
  if has_type {
    local filetype = lower(regexs(1))
    if `"`filetype'"' != ".docx" error 198 // invalid syntax. a la -16512
  }
  else {
    local using `"`using'.docx"'
  }

  scalar replace = "`replace'" == "replace"
  cap confirm file `"`using'"'
  if (_rc == 0 & !replace) error 602 // file already exists

  mata: st_global("word_handle", strofreal(_docx_new()))
  if ($word_handle == -16521) {
    cap mata: _docx_closeall()
    error 681 // too many open files. a la -16521
  }
  global word_filename `"`using'"'
  word_save
end

prog def word_save
  syntax [anything] [using/] [, replace]

  is_blank `"`anything'"'
  if `r(is_blank)' local anything = $word_handle

  is_blank `"`using'"'
  if `r(is_blank)' local using = `"$word_filename"'

  is_blank `"`replace'"'
  if `r(is_blank)' local replace = 1

  mata_func `"_docx_save(`anything', `"`using'"', `replace')"'
end


// helpers
// -------

prog def mata_func
  args func success failure
  mata: st_local("rc", strofreal(`func'))
  if `rc' == 0 {
    is_blank `"`success'"'
    if !r(is_blank) di as res `"`success'"'
  }
  else {
    local errmsg `"r(`rc'); `failure'"'
    di as err `"`errmsg'"'
    mata: _error(`"`errmsg'"')
  }
end

prog def has_handle, rclass
  scalar n = 0

  is_blank `"$word_handle"'
  if r(is_blank) scalar n = n + 1

  is_blank `"$word_filename"'
  if r(is_blank) scalar n = n + 1

  if n == 0 return scalar has_handle = 1
  else if n == 2 return scalar has_handle = 0
  else error 197 // invalid syntax. this shouldn't happen
end

prog def is_blank, rclass
  args s
  return scalar is_blank = trim(`"`s'"') == ""
end

prog def unquote, rclass
  args s
  scalar dquoted = regexm(`"`s'"', `"^\"(.*)\"$"')
  if dquoted return local unquoted = regexs(1)
  else {
    scalar cdquoted = regexm(`"`s'"', `"^\`\"(.*)\"\'$"')
    if cdquoted return local unquoted = regexs(1)
    else return local unquoted = `"`s'"'
  }
end

