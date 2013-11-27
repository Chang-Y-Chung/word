// certification script for word.ado

clear
discard
set more off

about
query compilenumber

// setup
// -----
// put word.ado, and test.png in the same directory as this file
//   and cd to this directory

// remove test.docx, if it exists
//   i.e. !del test.docx

// make sure that your ado path include the current directory
//   i.e. adopath

// run this
//   i.e. do word (in the command window of stata)
!rm -f test.docx

// test 1
// ------
cscript "open and close" adofiles word

word open using test.docx
assert $word_handle == 0
assert $word_filename == "test.docx"

word add image using test.png
word close
// open and see if test.docx has the image




// test 2
// ------
cap word open using test
di "_rc=" _rc
// should fail, since test.docx already exists

word open using test, replace
// should work

// clean up
word close


// test 3
// ------
word open using test, replace
forval i=1/6 {
  word add image using test.png, cx(1440) // in twips
}
word close
// test.docx should have tiny 6 pictures


// test 4
// ------
// the keyword using is now optional
word open test, replace
word add image test.png, cx(2880) // in twips
word add image test.png, cx(2880) // ditto
word close
// test.docx should have two pictures whose width is 2 inches


// test 5
// ------
// add text
word open test, replace
word add text test1, style("Title")
word add text test 2
word add text "test 3", style(Heading1)
word add text `"test 4"'
word add text ""
word add text `"test 6"' some more
word add text `"`"test 7"' some more"', s(Heading3)
word close
// test.docx should have 7 paragraphs
