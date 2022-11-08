Attribute VB_Name = "modAddBackSlash"
' #VBIDEUtils#************************************************************
' * Author           : Larry Rebich
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           : larry@buygold.net
' * Date             : 12/10/2003
' * Purpose          :
' * Project Name     : DBUpdateADO
' * Module Name      : modAddBackSlash
' **********************************************************************
' * Comments         :
' *
' *
' * Example          :
' *
' * History          : Updated by Waty Thierry
' * 2000/04/27 Copyright © 2000, Larry Rebich
' * 2000/04/27 larry@buygold.net, www.buygold.net, 760.771.4730
' * 2000/10/01 Used in BrandingModel
' * 2000/10/01 Used in Branding
' * 2003/01/31 Used in MBS
' *
' * See Also         :
' *
' *
' **********************************************************************

Option Explicit
DefLng A-Z

Const mcsBkSlash = "\"

Public Function AddBackslash(sThePath As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modAddBackSlash
   ' * Procedure Name   : AddBackslash
   ' * Parameters       :
   ' *                    sThePath As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' Add a backslash to a path if needed
   ' sPath contains the path
   ' Return a path with a backslash
   If Right$(sThePath, 1) <> mcsBkSlash Then
      sThePath = sThePath + mcsBkSlash
   End If
   AddBackslash = sThePath
End Function
