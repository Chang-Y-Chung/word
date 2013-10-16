
// tests for word.ado

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


// test 1
// ------
word open using test.docx
word add image using test.png
word close
// open and see if test.docx has the image


// test 2
// ------
cap word open using test
di "_rc=`rc'"
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

