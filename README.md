# Introduction

*Womacs* is a set of VBA macros to add *Emacs* style key bindings and commands to *MS Word*. It presents you an Emacs-like Word.


*Bring The Light of Emacs to MS Word Users.*

# Key bindings
## GNU Emacs like key bindings
Key binding|Command|Description
--- | --- | ---
    C-f|forward\_char|Move point right one character.
    C-b|backward\_char|Move point left one character.
    M-f|forward\_word|point forward one word.
    M-b|backward\_word|backward until encountering the beginning of a word.
    C-a|move\_beginning\_of\_line|Move point to beginning of current line.
    C-e|move\_end\_of\_line|Move point to end of current line.
    C-p|previous\_line|Move cursor vertically up one line.
    C-n|next\_line|Move cursor vertically down one line.
    C-d|delete\_char|Delete the following one character.
    M-d|kill\_word|Kill characters forward until encountering the end of a word.
    M-backspace|backward\_kill\_word|Kill characters backward until encountering the beginning of a word.
    C-SPC|set\_mark\_command|Set the mark where point is.
    C-@|set\_mark\_command|Set the mark where point is.
    C-g|keyboard\_quit|Signal a `quit' condiction.
    M-w|kill\_ring\_save|Copy
    C-w|kill\_region|Kill ("cut") text between point and mark.
    C-c w|delete\_region|Delete the text between point and mark.
    C-y|yank|Paste
    C-k|kill\_line|Kill the rest of the current line; if no nonblanks there, kill thru newline.
    M-u|upcase\_word|Convert following word to upper case, moving over.
    M-l|downcase\_word|Convert following word to lower case, moving over.
    M-c|capitalize\_word|Capitalize the following word (or ARG words), moving over.
    C-m|newline|Insert a newline.
    C-u|universal\_argument|Begin a numeric argument for the following command.
    M-x|execute\_extended\_command|Read function name, then read its arguments and call it.
    C-t|transpose\_chars|Interchange characters around point, moving forward one character.
    M-t|transpose\_words|Interchange words around point, leaving point at end of them.
    M-a|backward\_sentence|Move backward to start of sentence.
    M-e|forward\_sentence|Move forward to next end of sentence.
    C-x [|backward\_page|Move backward to page boundary.
    C-x ]|forward\_page|Move forward to page boundary.
    C-s|isearch\_forward|Do incremental search forward.
    C-r|isearch\_backward|Do incremental search backward.
    C-z|undo|Undo some previous changes.
    C-/|undo|Undo some previous changes.
    M-}|forward\_paragraph|Move forward to end of paragraph.
    M-{|backward\_paragraph|Move backward to start of paragraph.
    M-v|scroll\_down|Scroll text of selected window down ARG lines.
    C-v|scroll\_up|Scroll text of selected window upward ARG lines.
    C-l|recenter|Move current buffer line to the specified window line.
    C-o|open\_line|Insert a newline and leave point before it.
    C-home|beginning\_of\_buffer|Move point to the beginning of the buffer.
    C-end|end\_of\_buffer|Move point to the end of the buffer.
    C-x C-l|downcase\_region|Convert the region to lower case.
    C-x C-u|upcase\_region|Convert the region to upper case.
    C-x C-s|save\_buffer|Save current buffer in visited file if modified.
    C-x C-x|exchange\_point\_and\_mark|Put the mark where point is now, and point where the mark is now.
    C-x h|mark\_whole\_buffer|Put point at beginning and mark at end of buffer.
    C-x o|other\_window|Activate another pane.
    C-x 0|delete\_window|Close current pane.
    C-x 1|delete\_other\_windows|Remove the document window split.
    C-x 2|split\_window\_vertically|Split the document window.
    M-j l|set\_justification\_left|Align text to the left.
    M-j c|set\_justification\_center|Center text.
    M-j r|set\_justification\_right|Align text to the right.
    M-j b|set\_justification\_full|Justify.
    C-h k|describe\_key|Display documentation of the function invoked by KEY.
    C-h f|describe\_function|Display the documentation of FUNCTION (a symbol).
    M-\|delete\_horizontal\_space|Delete all spaces and tabs around point.


# MS word specific key bindings (these bindings are prone to change) 

| Tables        | Are           | Cool  |
| ------------- |:-------------:| -----:|
| col 3 is      | right-aligned | $1600 |
| col 2 is      | centered                |   $12 |
| zebra stripes | are neat      |    $1 |
