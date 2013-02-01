" outlook.vim - Edit emails using Vim from Outlook
" ---------------------------------------------------------------
" Version:       8.0
" Authors:       David Fishburn <dfishburn dot vim at gmail dot com>
" Last Modified: 2013 Jan 10
" Created:       2009 Jan 17
" Homepage:      http://www.vim.org/scripts/script.php?script_id=3087
" Help:          :h outlook.txt 
"


" Only use this on the Windows platforms
if !has("win32") && !has("win95")  && !has("win64") 
    finish
endif

if exists('g:loaded_outlook')
    finish
endif

" Turn on support for line continuations when creating the script
let s:cpo_save = &cpo
set cpo&vim

" Capture the output 
let g:outlook_save_cscript_output = 1

" View errors if updates fail
let g:outlook_view_cscript_error = 1

" Whether to open the email in a new tab 
if !exists('g:outlook_use_tabs')
    let g:outlook_use_tabs = 0
endif

" Whether to delete the buffer on save
if !exists('g:outlook_nobdelete')
    let g:outlook_nobdelete = 1
endif

" autoindent - preserve indent level on new line
if !exists('g:outlook_noautoindent')
    setlocal autoindent 
endif

" javascript  - location of the outlookvim.js file
if !exists('g:outlook_javascript')
    " Default location for the outlookvim.js file.
    " This can be overridden in your vimrc via
    " g:outlook_javascript
    " let g:outlook_javascript = expand('$VIM/vimfiles/plugin/outlookvim.js')
    let g:outlook_javascript = expand('<sfile>:p:h').'/outlookvim.js'
endif

" textwidth  - automatically wrap at a column
if exists('g:outlook_textwidth')
    let &textwidth = g:outlook_textwidth
else
    let &textwidth = 76
endif

" servername  - Choose which Vim server instance to edit the email with
if !exists('g:outlook_servername')
    let g:outlook_servername = ''
endif

" globpath uses wildignore and suffixes options unless a flag is provided
" to ignore those settings.
let s:ignore_wildignore_setting = 1

function! Outlook_WarningMsg(msg)
    echohl WarningMsg
    echomsg "outlook: ".a:msg 
    echohl None
endfunction
      
function! Outlook_ErrorMsg(msg)
    echohl ErrorMsg
    echomsg "outlook: ".a:msg 
    echohl None
endfunction

function! Outlook_ExecuteJS(filename)
    let cmd = "let result = system('cscript \""
    let cmd = cmd . expand(g:outlook_javascript)
    let cmd = cmd . '" "'.a:filename.'"'
    let cmd = cmd . "  ')"
    exec cmd

    echomsg result
endfunction

function! Outlook_EditFile(filename, encoding)
    if g:outlook_servername == '' || g:outlook_servername == v:servername 
        let remove_bufnr = -1
        if !filereadable(a:filename) 
            call Outlook_ErrorMsg( 'Cannot find filename['.a:filename.']')
            return
        endif

        if bufname('%') == '' && winnr('$') == 1 && &modified != 1
            let remove_bufnr = bufnr('%')
            let cmd = ':'.
                        \ 'new'.
                        \ (a:encoding==''?'':' ++enc='.a:encoding).
                        \ ' '.a:filename.
                        \ "\n"
        else
            let cmd = ':'.
                        \ (g:outlook_use_tabs == 1 ? 'tabnew' : 'e').' '.
                        \ (a:encoding==''?'':' ++enc='.a:encoding).
                        \ ' '.a:filename.
                        \ "\n"
        endif
        let g:outlook_last_cmd = cmd
        exec cmd

        if remove_bufnr != -1
            exec remove_bufnr.'bw'
        endif
    else
        if match( serverlist(), '\<'.g:outlook_servername.'\>' ) > -1
            call remote_send( g:outlook_servername, ":call Outlook_EditFile( '".a:filename."', '".a:encoding."' )\<CR>" )
            " call remote_send( g:outlook_servername, cmd )
        else
            call Outlook_ErrorMsg( 'Please start a new Vim instance using "gvim --servername '.g:outlook_servername.'"')
        endif
    endif
endfunction

function! Outlook_BufWritePost()
    " autoread, prevents this message:
    "      "File has changed, do you want to Load" 
    if !exists('g:outlook_noautoread')
        setlocal autoread 
    endif

    let filename = expand("<afile>:p")
    if filename == ''
        let filename = expand("%:p")
        call s:Outlook_WarningMsg( 'Filename was blank, using:'.filename )
    endif

    let cmd = 'cscript "'. expand(g:outlook_javascript). 
                \ '" "'.
                \ filename.
                \ '" '.
                \ g:outlook_nobdelete

    try
        let g:outlook_cscript_output = system(cmd)
    catch /.*/
        call Outlook_WarningMsg( 'Outlook_BufWritePost: Error calling cscript to update Outlook:'.v:exception )
        if g:outlook_save_cscript_output == 1 && g:outlook_view_cscript_error
           call Outlook_WarningMsg( v:exception )
       endif
    finally
    endtry

    if g:outlook_save_cscript_output == 1 && g:outlook_view_cscript_error
        if g:outlook_cscript_output =~ '\c\(OutlookVim\[\d\+\]:\s*\(Successfully\)\@<=\|runtime error\)' 
           call Outlook_WarningMsg( substitute(g:outlook_cscript_output, '\c^.*\(OutlookVim\[.*\)', '\1', '') )
        elseif g:outlook_nobdelete == 0 
            bdelete 
        endif
    else
        if g:outlook_nobdelete == 0 
            bdelete 
        endif 
    endif
   
endfunction

function! Outlook_BufUnload()
    " This is necessary in case the buffer is not saved (:w), since
    " the temporary files are (by default) deleted by outlookvim.js.
    " But the javascript is only called if you save the buffer, if
    " you choose to abandon your edits (:bw!), the temporary files are
    " left over in the temporary directory.
    if !exists('g:outlook_nodelete_unload') || g:outlook_nodelete_unload != 1 
        if expand('<afile>:e') == 'outlook'
            call delete(expand('<afile>:p').'.ctl')
            call delete(expand('<afile>:p'))
        endif
    endif
endfunction

" Check is cscript.exe is already in the PATH
" Some path entries may end in a \ (c:\util\;), this must also be replaced
" or globpath fails to parse the directories
if v:version > 703 || v:version == 702 && has('patch051')
    let cscript_location = globpath(substitute($PATH, '\\\?;', ',', 'g'), 'cscript.exe', s:ignore_wildignore_setting)
else
    let cscript_location = globpath(substitute($PATH, '\\\?;', ',', 'g'), 'cscript.exe')
endif
if strlen(cscript_location) == 0
    call Outlook_ErrorMsg("Cannot find cscript.exe in system path")
    finish
endif
    
if exists('g:outlook_javascript')
    if !filereadable(expand(g:outlook_javascript)) 
        call Outlook_ErrorMsg("Cannot find javascript file[" .
                \ expand(g:outlook_javascript) .
                \ ']' )
        finish
    endif
else
    call Outlook_ErrorMsg("Cannot find the variable: g:outlook_javascript ")
    finish
endif

" These autocommands only need to be created once,
" so store a script variable to prevent reloading
if has('autocmd') && !exists("g:loaded_outlook")
    " Save the current default register
    let saveB = @"

    " Create a group of the autocommands for outlook
    augroup outlook
    " Remove the previous group (if it existed)
    au!

    " Each time we enter this buffer, set the filetype to mail
    " which allows us to rely on each personal users mail preferences
    autocmd BufEnter *.outlook setlocal filetype=mail

    " nested is required since we are issuing a bdelete, inside an autocmd
    " so we also need the required autocmd to fire for that command.
    " setlocal autoread, prevents this message:
    "      "File has changed, do you want to Load" 
    autocmd BufWritePost *.outlook nested call Outlook_BufWritePost()

    autocmd BufUnload *.outlook call Outlook_BufUnload()

    augroup END
    
    " Don't re-run the script if already sourced
    let g:loaded_outlook = 8

    let @"=saveB
endif

let &cpo = s:cpo_save
unlet s:cpo_save

" vim:fdm=marker:nowrap:ts=4:expandtab:
