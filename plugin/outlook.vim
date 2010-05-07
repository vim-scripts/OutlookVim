" outlook.vim - Edit emails using Vim from Outlook
" ---------------------------------------------------------------
" Version:       2.0
" Authors:       David Fishburn <dfishburn dot vim at gmail dot com>
" Last Modified: 2010 May 06
" Created:       2009 Jan 17
" Homepage:      http://vim.sourceforge.net/script.php?script_id=???
" Help:         :h outlook.txt 
"


" Only use this on the Windows platforms
if !has("win32") && !has("win95")  && !has("win64") 
    finish
endif

if exists('g:loaded_outlook') || &cp
    finish
endif

" Capture the output 
let g:outlook_save_cscript_output = 1

" View errors if updates fail
let g:outlook_view_cscript_error = 1

" Default location for the outlookvim.js file.
" This can be overridden in your vimrc via
" g:outlook_javascript
let s:outlook_javascript_default = expand('$VIM/vimfiles/plugin/outlookvim.js')

" autoindent - preserve indent level on new line
if !exists('g:outlook_noautoindent')
    setlocal autoindent 
endif

" textwidth  - automatically wrap at a column
if exists('g:outlook_textwidth')
    let &textwidth = g:outlook_textwidth
else
    let &textwidth = 76
endif

function! Outlook_WarningMsg(msg)
    echohl WarningMsg
    echomsg a:msg 
    echohl None
endfunction
      
function! Outlook_ErrorMsg(msg)
    echohl ErrorMsg
    echomsg a:msg 
    echohl None
endfunction

function! Outlook_ExecuteJS(filename)
    let cmd = "let result = system('cscript \""
    if exists('g:outlook_javascript')
        let cmd = cmd . expand(g:outlook_javascript)
    else
        let cmd = cmd . s:outlook_javascript_default
    endif
    let cmd = cmd . '" "'.a:filename.'"'
    let cmd = cmd . "  ')"
    exec cmd

    echomsg result
endfunction

" Check is cscript.exe is already in the PATH
" Some path entries may end in a \ (c:\util\;), this must also be replaced
" or globpath fails to parse the directories
if strlen(globpath(substitute($PATH, '\\\?;', ',', 'g'), 'cscript.exe')) == 0
    call Outlook_ErrorMsg("outlook: Cannot find cscript.exe in system path")
    finish
endif
    
if exists('g:outlook_javascript')
    if !filereadable(expand(g:outlook_javascript)) 
        call Outlook_ErrorMsg("outlook: Cannot find javascript: " .
                \ expand(g:outlook_javascript) )
        finish
    endif
else
    if !filereadable(expand(s:outlook_javascript_default)) 
        call Outlook_ErrorMsg("outlook: Cannot find javascript: " .
                \ expand(s:outlook_javascript_default) )
        finish
    endif
endif

" These autocommands only need to be created once,
" so store a script variable to prevent reloading
if has('autocmd') && !exists("g:loaded_outlook")
    " Save the current default register
    let saveB = @"

    " Check to see if the BufWritePost autocommand already exists
    redir @"
    silent! exec 'augroup'
    redir END

    if @" !~? '\<outlook\>' 
    
        " Create a group of the autocommands for outlook
        augroup outlook
        " Remove the previous group (if it existed)
        au!

        " Each time we enter this buffer, set the filetype to mail
        " which allows us to rely on each personal users mail preferences
        let cmd = 'autocmd BufEnter *.outlook setlocal filetype=mail '
        exec cmd

        " nested is required since we are issuing a bdelete, inside an autocmd
        " so we also need the required autocmd to fire for that command.
        " setlocal autoread, prevents this message:
        "      "File has changed, do you want to Load" 
        
        let cmd = 'autocmd BufWritePost *.outlook nested '
        if !exists('g:outlook_noautoread')
            let cmd = cmd . 'setlocal autoread | '
        endif

        " silent! !start is to prevent the "Press any key to continue message"
        " This is behaviour I (dfishburn) prefer, when we write the file, the
        " buffer is deleted, and removed from the list of files in the Vim
        " session.  outlookvim.js has already deleted the temporary file from the
        " filesystem.
        let cmd = cmd . "let g:outlook_cscript_output = system('cscript \""
        if exists('g:outlook_javascript')
            let cmd = cmd . expand(g:outlook_javascript)
        else
            let cmd = cmd . s:outlook_javascript_default
        endif
        let cmd = cmd . '" "'."'".'.expand("%").'."'".'"'
        let cmd = cmd . "  ')"

        if !exists('g:outlook_nobdelete')
            let cmd = cmd . "| bdelete "
        endif
        if g:outlook_save_cscript_output == 1 && g:outlook_view_cscript_error
            let cmd = cmd . "| if g:outlook_cscript_output =~ 'outlookvim:' | call Outlook_WarningMsg( substitute(g:outlook_cscript_output, '^.*\\(outlookvim:.*\\)', '\\1', '') ) | endif "
        endif
       
        exec cmd

        " This is necessary in case the buffer is not saved (:w), since
        " the temporary files are only deleted by outlookvim.js.
        " But the javascript is only called if you save the buffer, if
        " you choose to abandon your edits, the temporary files are left
        " over in the temporary directory.
        if !exists('g:outlook_nodelete_unload')
            let cmd =  "autocmd BufUnload *.outlook " .
                        \ "if expand('%:e') == 'outlook' " .
                        \ "| call delete(expand('%:p').'.ctl') "  .
                        \ "| call delete(expand('%:p')) " .
                        \ "| endif"
            exec cmd
        endif

        augroup END
    
    endif

    " Don't re-run the script if already sourced
    let g:loaded_outlook = 1

    let @"=saveB
endif

" vim:fdm=marker:nowrap:ts=4:expandtab:
