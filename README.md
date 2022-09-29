# excel_vba_funcReduceSpaces
A quick function to correct extraneous spaces from sloppy typers.

I used to deal with sloppy typers quite a bit. I had one lady who was always hitting
the space bar in the weirdest places. This wound up causing a great deal of problems
with procedures that required specific formatting of input (it was a technical environment
so I was even doubly frustrated because she knew it was a problem). So this is what
I came up to deal with it. It goes back to the old programmers adage:

**Never trust users to enter what you expect them to enter.**

This procedure will accept a string and return a string. It looks for multiple
contiguous spaces and then reduces them to a single space.

For example:

strInput = "This is a sentence  with an &nbsp;extra &nbsp;&nbsp;space &nbsp;&nbsp;&nbsp;or three."

Using <code>funcReduceSpaces(strInput)</code> will return:

"This is a sentence with an extra space or three."

I know that popular belief is that using Trim(strInput) will remove extra spaces. It will
remove them but only at the beginning or end. It does not remove them mid string.

NOTE: This will be used with user input that is immediately processed by VBA. You won't
be able to use this in a forumla unless the formula is grabbing user input from another cell.

Download the example workbook and look in the VBA editor, there are a couple testing
subroutines in there. You can copy and paste the code into your project or import
the replace_spaces.bas in the repository.
