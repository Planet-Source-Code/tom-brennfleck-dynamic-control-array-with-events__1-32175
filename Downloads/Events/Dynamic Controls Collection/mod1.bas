Attribute VB_Name = "mod1"
Option Explicit

'   Copyright T. Brennfleck SamSa Consulting Group Pty. Ltd. 2001
'   This code is copyrighted and has limited warranties.
'   Terms of Agreement:
'   By using this code, you agree to the following terms...
'   1) You may use this code in your own programs
'      (and may compile it into a program and distribute it in compiled
'      format for langauges that allow it) freely and with no charge.
'   2) You MAY NOT redistribute this code (for example to a web site, or code library, or book)
'      without written permission from the original author.
'      Failure to do so is a violation of copyright laws.
'   3) If you use this code for a commercial product I would like to have an email
'      stating the product that it is being used in.

Public EventHandler As cEventHandler

Sub main()

    Set EventHandler = New cEventHandler
    frm1.Show
    
End Sub
