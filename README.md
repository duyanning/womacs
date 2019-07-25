# Introduction

**Womacs** is a set of VBA macros to add **Emacs** style key bindings and commands to **MS Word**. It presents you an Emacs-like Word.


**Bring The Light of Emacs to MS Word Users.**

# Key bindings
## GNU Emacs like key bindings
```
Key binding 	Command 	Description 
C-f 	forward_char 	Move point right one character.
C-b 	backward_char 	Move point left one character.
M-f 	forward_word 	point forward one word.
M-b 	backward_word 	backward until encountering the beginning of a word.
C-a 	move_beginning_of_line 	Move point to beginning of current line.
C-e 	move_end_of_line 	Move point to end of current line.
C-p 	previous_line 	Move cursor vertically up one line.
C-n 	next_line 	Move cursor vertically down one line.
C-d 	delete_char 	Delete the following one character.
M-d 	kill_word 	Kill characters forward until encountering the end of a word.
M-backspace 	backward_kill_word 	Kill characters backward until encountering the beginning of a word.
C-SPC 	set_mark_command 	Set the mark where point is.
C-@ 	set_mark_command 	Set the mark where point is.
C-g 	keyboard_quit 	Signal a `quit' condiction.
M-w 	kill_ring_save 	Copy
C-w 	kill_region 	Kill ("cut") text between point and mark.
C-c w 	delete_region 	Delete the text between point and mark.
C-y 	yank 	Paste
C-k 	kill_line 	Kill the rest of the current line; if no nonblanks there, kill thru newline.
M-u 	upcase_word 	Convert following word to upper case, moving over.
M-l 	downcase_word 	Convert following word to lower case, moving over.
M-c 	capitalize_word 	Capitalize the following word (or ARG words), moving over.
C-m 	newline 	Insert a newline.
C-u 	universal_argument 	Begin a numeric argument for the following command.
M-x 	execute_extended_command 	Read function name, then read its arguments and call it.
C-t 	transpose_chars 	Interchange characters around point, moving forward one character.
M-t 	transpose_words 	Interchange words around point, leaving point at end of them.
M-a 	backward_sentence 	Move backward to start of sentence.
M-e 	forward_sentence 	Move forward to next end of sentence.
C-x [ 	backward_page 	Move backward to page boundary.
C-x ] 	forward_page 	Move forward to page boundary.
C-s 	isearch_forward 	Do incremental search forward.
C-r 	isearch_backward 	Do incremental search backward.
C-z 	undo 	Undo some previous changes.
C-/ 	undo 	Undo some previous changes.
M-} 	forward_paragraph 	Move forward to end of paragraph.
M-{ 	backward_paragraph 	Move backward to start of paragraph.
M-v 	scroll_down 	Scroll text of selected window down ARG lines.
C-v 	scroll_up 	Scroll text of selected window upward ARG lines.
C-l 	recenter 	Move current buffer line to the specified window line.
C-o 	open_line 	Insert a newline and leave point before it.
C-home 	beginning_of_buffer 	Move point to the beginning of the buffer.
C-end 	end_of_buffer 	Move point to the end of the buffer.
C-x C-l 	downcase_region 	Convert the region to lower case.
C-x C-u 	upcase_region 	Convert the region to upper case.
C-x C-s 	save_buffer 	Save current buffer in visited file if modified.
C-x C-x 	exchange_point_and_mark 	Put the mark where point is now, and point where the mark is now.
C-x h 	mark_whole_buffer 	Put point at beginning and mark at end of buffer.
C-x o 	other_window 	Activate another pane.
C-x 0 	delete_window 	Close current pane.
C-x 1 	delete_other_windows 	Remove the document window split.
C-x 2 	split_window_vertically 	Split the document window.
M-j l 	set_justification_left 	Align text to the left.
M-j c 	set_justification_center 	Center text.
M-j r 	set_justification_right 	Align text to the right.
M-j b 	set_justification_full 	Justify.
C-h k 	describe_key 	Display documentation of the function invoked by KEY.
C-h f 	describe_function 	Display the documentation of FUNCTION (a symbol).
M-\ 	delete_horizontal_space 	Delete all spaces and tabs around point.
    
```

# MS word specific key bindings (these bindings are prone to change) 

```
Key binding 	Function
C-c C-y 	Paste as plain text.
C-c C-c 	Toggle displaying all nonprinting characters.
C-c b 	Toggle bold
C-c i 	Toggle italic
C-c C-u 	Toggle superscript, such as X2
C-c C-l 	Toggle subscript, such as X2
C-c x 	Toggle strike through
C-c h 	Toggle highlight (yellow)
C-c r 	Toggle font color (red)
C-c 1 	Style Heading 1
C-c 2 	Style Heading 2
C-c 3 	Style Heading 3
C-c 4 	Style Heading 4
C-c 5 	Style Heading 5
C-c n 	Style Normal
C-c d 	Style Normal Indent
C-c t 	Style Title
C-c [ ' 	Insert Single Opening Quote
C-c ] ' 	Insert Single Closing Quote
C-c [ " 	Insert Double Opening Quote
C-c ] " 	Insert Double Closing Quote
```

# Customization

right click womacs.dotm and choose open.

press Alt+F11 to get the VBA editor.

double click module dotEmacs.

you can use `global_set_key`/`global_unset_key` here.


# Download

https://github.com/duyanning/womacs/releases

# Install

```Batchfile
copy womacs.dotm "%userprofile%\Application Data\Microsoft\Word\Startup"
```

Caution: Before installing, you need to exit all MS-word. Perhaps, you may want to backup up the dotEmacs module of the old version.
# Activate

run macro `toggle_womacs` to toggle womacs. 
It is handy to add a Quick Access Toolbar button for this macro (`Project.Main.toggle_womacs`).
# Bug report

I have been developing Womacs using MS-Word 2010 and Win7 (Chinese Edition).

If you experienced problems, please visit https://github.com/duyanning/womacs/issues and click `New issue` to submit bug reports.

Please include in bug reports the version number of Womacs in which you are running into troubles.

You can get the version nubmer by M-x version.

If you get an error like "Compile error in hidden module: ImeHelper" when invoking toggle_womacs, please download the latest version and follow steps:

right click and open womacs.dotm.

press Alt-F11 to get the VBA IDE.

At the top the ImeHelper module,

change

has_imm = 1

to

has_imm = 0

then save and reinstall.

If it does not work, please send me an email: duyanning at gmail com

note: The ImeHelper module is used to manipulate Input Method Editor which is common on Windows for East Asian. If you do not use IME, you can change 'has_imm' to 0.


# Acknowledgments

I copied a few of code from another GPL-licensed free software VBAcs (http://rath.ca/Misc/VBacs/).

However, I learned a lot of knowledge about VBA and Word Object Model from VBAcs.

I am very grateful to the two authors of VBAcs: Christopher Rath & John Cooper.

Thank you!!!
# Handy tools

In addition to imitating Emacs, Womacs also offers some useful VBA macros. 
