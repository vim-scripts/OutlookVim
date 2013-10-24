// outlookvim.js
//
// Author:        David Fishburn
// Version:       13.0
// Last Modified: 2013 Sep 25
// Homepage:      http://www.vim.org/scripts/script.php?script_id=3087
//
// Purpose:
//   To be used in conjunction with the OutlookVim plugin to allow
//   Vim to update an Outlook email which has been edited using
//   Vim.  Saving the file in Vim will automatically trigger this
//   Javascript file to be called which uses Outlooks APIs to
//   update the body of the email message in Outlook.
//
//   This has been tested with Outlook version 2003.
//
// Reference:
//   Overview of Windows Scripting
//       http://msdn2.microsoft.com/en-us/library/ms950396.aspx
//
//   Microsoft JScript Documentation
//       http://msdn2.microsoft.com/en-us/library/hbxc2t98.aspx
//
//       JScript Language Reference
//           http://msdn2.microsoft.com/en-us/library/yek4tbz0.aspx
//
//   JavaScript API for Office
//       http://msdn.microsoft.com/en-us/library/fp160953.aspx
//
//       Example
//           http://stackoverflow.com/questions/12247696/open-outlook-with-javascript-for-sending-mail-with-attachment-in-c-sharp

var objArgs     = WScript.Arguments;
var version     = 13;

function updateOutlook( emailfile, persistfiles, bodyformat )
{
    var ctlfile     = emailfile + ".ctl";
    var outlook     = null;
    var fs          = null;
    var f           = null;
    var fid         = null;
    var objNS       = null;
    var objInbox    = null;
    var entryID     = null;
    var msg         = null;
    var inspector   = null;
    var readOnly    = 1;
    var createNo    = false;
    var mixedMode   = -2;
    var updateMsg   = true;

    WScript.Echo("OutlookVim[" + version + "]: updateOutlook persist files:" + persistfiles);
    try
    {
	outlook     = new ActiveXObject("Outlook.Application");
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Unable to create Outlook.Application:"+err.message);
	return;
    }
    try
    {
	objNS       = outlook.GetNamespace("MAPI");
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Unable to get outlook namespace["+err.message+"]");
	return;
    }
    try
    {
	fs          = new ActiveXObject("Scripting.FileSystemObject");
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Unable to create Scripting.FileSystemObject:"+err.message);
	return;
    }
    try
    {
        // Allow the file to be opened in unicode as well (mixedMode)
        // Parameters
        //     Filename
        //     IOMode
        //           1 = ForReading
        //           2 = ForWriting
        //           8 = ForAppending
        //     Create (Boolean)
        //     Format
        //            0 = TristateFalse: Open the file as ASCII (Default value)
        //           -1 = TristateTrue: Open the file as Unicode
        //           -2 = TristateMixed: Mixed mode
        //           -2 = TristateUseDefault: Open the file as System Default type.
	f = fs.OpenTextFile(emailfile, readOnly, createNo, mixedMode);
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Unable to open file["+emailfile+"] Err["+err.message+"]");
	return;
    }
    try
    {
	objInbox    = objNS.GetDefaultFolder(6);
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Unable to get Inbox["+err.message+"]");
	return;
    }
    try
    {
	fid         = fs.OpenTextFile(ctlfile);
    }
    catch(err)
    {
        updateMsg = false;
	WScript.Echo("OutlookVim[" + version + "]: Unable to open control file["+ctlfile+"] assuming new email Err["+err.message+"]");
	//return;
    }
    if( updateMsg )
    {
        try
        {
            entryID     = fid.ReadLine();
        }
        catch(err)
        {
            WScript.Echo("OutlookVim[" + version + "]: Failed to read control file["+ctlfile+"] Err["+err.message+"]");
            return;
        }
        try
        {
            msg      = objNS.GetItemFromID(entryID);
        }
        catch(err)
        {
            WScript.Echo("OutlookVim[" + version + "]: GetItemFromID failed["+err.message+"]");
            return;
        }
    }
    else
    {
        // value 0 = MailItem
        msg = outlook.CreateItem(0);
    }

    try
    {
        switch( bodyformat )
        {
        case "html":
        case "HTML":
            msg.HTMLBody = f.ReadAll();
            break;
        //case "rtf":
        //case "RTF":
        //    msg.RTFBody = f.ReadAll();
        //    break;
        default:
            msg.Body = f.ReadAll();
            break;
        }
    }
    catch(err)
    {
        WScript.Echo("OutlookVim[" + version + "]: Failed to read email file["+emailfile+"] Err["+err.message+"]");
        return;
    }

    if( updateMsg )
    {
        fid.Close();
        f.Close();
    }

    try
    {
	inspector = msg.GetInspector;
	inspector.Activate();
    }
    catch(err)
    {
	WScript.Echo("OutlookVim[" + version + "]: Failed to get Inspector["+err.message+"]");
	return;
    }

    if( updateMsg )
    {
        if( 1 == persistfiles )
        {
            WScript.Echo("OutlookVim[" + version + "]: Keeping files");
        }
        else
        {
            WScript.Echo("OutlookVim[" + version + "]: Deleting files[" + persistfiles + "]");
            try
            {
                f = fs.GetFile(emailfile);
                f.Delete();
            }
            catch(err)
            {
                WScript.Echo("OutlookVim[" + version + "]: Failed to get and delete email file["+emailfile+"] Err["+err.message+"]");
            }

            try
            {
                fid = fs.GetFile(ctlfile);
                fid.Delete();
            }
            catch(err)
            {
                WScript.Echo("OutlookVim[" + version + "]: Failed to get and delete control file["+ctlfile+"] Err["+err.message+"]");
            }
        }
    }

    WScript.Echo("OutlookVim[" + version + "]: Successfully updated Outlook, message ID["+entryID+"]");
}

if( 0 == objArgs.length )
{
    WScript.Echo("OutlookVim[" + version + "]: Hello from OutlookVim!");
} else {
    var emailfile = objArgs(0);
    var persistfiles = 0;
    var bodyformat = "plain";
    if( objArgs.length > 1 )
    {
        var persistfiles = objArgs(1);;
        WScript.Echo("OutlookVim[" + version + "]: Persist files, overriding to[" + persistfiles + "]");
    }
    if( objArgs.length > 2 )
    {
        var bodyformat = objArgs(2);;
        WScript.Echo("OutlookVim[" + version + "]: Body format overriding to[" + bodyformat + "]");
    }
    updateOutlook( emailfile, persistfiles, bodyformat );
}

