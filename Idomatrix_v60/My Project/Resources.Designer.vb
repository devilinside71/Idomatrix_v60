﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Idomatrix_v60.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap.
        '''</summary>
        Friend ReadOnly Property Idomatrix_32() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Idomatrix_32", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap.
        '''</summary>
        Friend ReadOnly Property Idomatrix_Delete_24() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Idomatrix_Delete_24", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap.
        '''</summary>
        Friend ReadOnly Property Idomatrix_Delete_32() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Idomatrix_Delete_32", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- end ReportEvalTable --&gt;
        '''
        '''&lt;!-- start ReportBodyEnd --&gt;
        '''&lt;/body&gt;
        '''&lt;!-- end ReportBodyEnd --&gt;
        '''
        '''&lt;/html&gt;.
        '''</summary>
        Friend ReadOnly Property ReportBodyEnd() As String
            Get
                Return ResourceManager.GetString("ReportBodyEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportEvalTable --&gt;
        '''&lt;table cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; class=&quot;evalopt01&quot; style=&quot;width: 40%&quot;&gt;
        '''	&lt;tr&gt;
        '''		&lt;td colspan=&quot;4&quot;&gt;Optimálishoz közeli eloszlás:&lt;/td&gt;
        '''	&lt;/tr&gt;
        '''	&lt;tr&gt;
        '''		&lt;td class=&quot;evalopt02&quot; style=&quot;width: 64%&quot;&gt;64%&lt;/td&gt;
        '''		&lt;td class=&quot;evalopt03&quot; style=&quot;width: 12%&quot;&gt;12%&lt;/td&gt;
        '''		&lt;td class=&quot;evalopt04&quot; style=&quot;width: 12%&quot;&gt;12%&lt;/td&gt;
        '''		&lt;td class=&quot;evalopt05&quot; style=&quot;width: 12%&quot;&gt;12%&lt;/td&gt;
        '''	&lt;/tr&gt;
        '''&lt;/table&gt;
        '''&lt;table cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; class=&quot;evalreal01&quot; style=&quot;width: 40%&quot;&gt;
        '''	&lt;tr&gt;
        '''		&lt;td cols [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportEvalTable() As String
            Get
                Return ResourceManager.GetString("ReportEvalTable", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportIntro --&gt;
        '''&lt;p class=&quot;head01&quot;&gt;&lt;strong&gt;&lt;span class=&quot;head02&quot;&gt;Időintervallum:&lt;/span&gt;&lt;/strong&gt;&lt;span class=&quot;head02&quot;&gt; 
        '''[INTERVAL]&lt;/span&gt;&lt;/p&gt;
        '''&lt;p class=&quot;head01&quot;&gt;&amp;nbsp;&lt;/p&gt;
        '''&lt;!-- end ReportIntro --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportIntro() As String
            Get
                Return ResourceManager.GetString("ReportIntro", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableEnd --&gt;
        '''&lt;/table&gt;
        '''&lt;p&gt;&amp;nbsp;&lt;/p&gt;
        '''&lt;!-- end ReportMainTableEnd --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableEnd() As String
            Get
                Return ResourceManager.GetString("ReportMainTableEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableFEnd --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m07&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;&lt;strong&gt;[LEFT1]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;&lt;strong&gt;[LEFT2]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;&amp;nbsp;&lt;/td&gt;
        '''&lt;td class=&quot;m08&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;&lt;strong&gt;[RIGHT1]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;&lt;strong&gt;[RIGHT2]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;&amp;nbsp;&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;!-- end ReportMainTableFEnd --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableFEnd() As String
            Get
                Return ResourceManager.GetString("ReportMainTableFEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableFRow --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m04&quot;&gt;[LEFT1]&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;[LEFT2]&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;[LEFT3]&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;[LEFT4]&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;[LEFT5]&lt;/td&gt;
        '''&lt;td class=&quot;m05&quot;&gt;[LEFT6]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT1]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT2]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT3]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT4]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT5]&lt;/td&gt;
        '''&lt;td class=&quot;m06&quot;&gt;[RIGHT6]&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;!-- end ReportMainTableFRow --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableFRow() As String
            Get
                Return ResourceManager.GetString("ReportMainTableFRow", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableFStart --&gt;
        '''&lt;table  cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; width=&quot;80%&quot;&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m01&quot; rowspan=&quot;[ROWSPAN]&quot;&gt;&lt;strong&gt;F&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;N&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;T&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''S&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m02&quot; colspan=&quot;6&quot;&gt;SÜRGŐS&lt;/td&gt;
        '''&lt;td class=&quot;m03&quot; rowspan=&quot;[ROWSPAN]&quot;&gt;F&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;N&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;T&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''S&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m03&quot; colspan=&quot; [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableFStart() As String
            Get
                Return ResourceManager.GetString("ReportMainTableFStart", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableNFEnd --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m15&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;&lt;strong&gt;[LEFT1]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;&lt;strong&gt;[LEFT2]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;&amp;nbsp;&lt;/td&gt;
        '''&lt;td class=&quot;m16&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;&lt;strong&gt;[RIGHT1]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;&lt;strong&gt;[RIGHT2]&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;&amp;nbsp;&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;!-- end ReportMainTableNFEnd --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableNFEnd() As String
            Get
                Return ResourceManager.GetString("ReportMainTableNFEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableNFRow --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m12&quot;&gt;[LEFT1]&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;[LEFT2]&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;[LEFT3]&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;[LEFT4]&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;[LEFT5]&lt;/td&gt;
        '''&lt;td class=&quot;m13&quot;&gt;[LEFT6]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT1]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT2]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT3]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT4]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT5]&lt;/td&gt;
        '''&lt;td class=&quot;m14&quot;&gt;[RIGHT6]&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;!-- end ReportMainTableNFRow --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableNFRow() As String
            Get
                Return ResourceManager.GetString("ReportMainTableNFRow", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMainTableNFStart --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;m09&quot; rowspan=&quot;[ROWSPAN]&quot;&gt;&lt;strong&gt;N&lt;br /&gt;
        '''&lt;/strong&gt;E&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;M&lt;strong&gt;&lt;br /&gt;
        '''&lt;br /&gt;
        '''&lt;/strong&gt;F&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;N&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;T&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''S&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;m10&quot; colspan=&quot;6&quot;&gt;SÜRGŐS&lt;/td&gt;
        '''&lt;td class=&quot;m11&quot; rowspan=&quot;[ROWSPAN]&quot;&gt;N&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;E&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;M&lt;strong&gt;&lt;br /&gt;
        '''&lt;br /&gt;
        '''&lt;/strong&gt;F&lt;strong&gt;&lt;br /&gt;
        '''&lt;/strong&gt;O&lt;strong&gt;&lt;br /&gt;
        '''&lt; [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportMainTableNFStart() As String
            Get
                Return ResourceManager.GetString("ReportMainTableNFStart", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMonthlyTableEnd  --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;auto-style9&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;alma&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;alma&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;auto-style10&quot; colspan=&quot;3&quot;&gt;&lt;strong&gt;ÖSSZ:&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;&lt;strong&gt;alma&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;&lt;strong&gt;alma&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;&lt;strong&gt;&lt;/strong&gt;&lt;/td&gt;
        '''&lt;/tr&gt;
        '''
        '''&lt;/table&gt;
        '''&lt;p&gt;&amp;nbsp;&lt;/p&gt;
        '''&lt;!-- end ReportMonthlyTableE [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportMonthlyTableEnd() As String
            Get
                Return ResourceManager.GetString("ReportMonthlyTableEnd", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMonthlyTableRow --&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;havi03&quot;&gt;[LEFT1]&lt;/td&gt;
        '''&lt;td class=&quot;havi04&quot;&gt;[LEFT2]&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;[LEFT3]&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;[LEFT4]&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;[LEFT5]&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;[LEFT6]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT1]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT2]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT3]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT4]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT5]&lt;/td&gt;
        '''&lt;td class=&quot;havi06&quot;&gt;[RIGHT6]&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;!-- end ReportMonthlyTableRow --&gt;.
        '''</summary>
        Friend ReadOnly Property ReportMonthlyTableRow() As String
            Get
                Return ResourceManager.GetString("ReportMonthlyTableRow", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportMonthlyTableStart --&gt;
        '''&lt;table  cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; width=&quot;80%&quot;&gt;
        '''&lt;tr&gt;
        '''&lt;td class=&quot;havi01&quot; colspan=&quot;6&quot;&gt;&lt;strong&gt;Havi célok&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi02&quot; colspan=&quot;6&quot;&gt;&lt;strong&gt;Havi feladatok&lt;/strong&gt;&lt;/td&gt;
        '''&lt;/tr&gt;
        '''
        '''&lt;tr&gt;
        '''&lt;td class=&quot;havi03&quot;&gt;&lt;strong&gt;T&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi04&quot;&gt;&lt;strong&gt;Megnevezés&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;Dátum&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;Terv&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;strong&gt;Tény&lt;/strong&gt;&lt;/td&gt;
        '''&lt;td class=&quot;havi05&quot;&gt;&lt;stro [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportMonthlyTableStart() As String
            Get
                Return ResourceManager.GetString("ReportMonthlyTableStart", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD XHTML 1.0 Transitional//EN&quot; &quot;http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd&quot;&gt;
        '''&lt;html xmlns=&quot;http://www.w3.org/1999/xhtml&quot;&gt;
        '''
        '''&lt;!-- start ReportPageHead --&gt;
        '''&lt;head&gt;
        '''&lt;meta content=&quot;text/html; charset=utf-8&quot; http-equiv=&quot;Content-Type&quot; /&gt;
        '''&lt;title&gt;Időmátrix&lt;/title&gt;
        '''&lt;style type=&quot;text/css&quot;&gt;
        '''.head01 {
        '''	font-family: Arial, Helvetica, sans-serif;
        '''}
        '''.head02 {
        '''	font-size: smaller;
        '''}
        '''
        '''
        '''.havi01 {
        '''	border-left: .5pt solid windowtext;
        '''	border-top: .5pt solid windowtex [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportStart() As String
            Get
                Return ResourceManager.GetString("ReportStart", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;!-- start ReportSumTable --&gt;
        '''&lt;table  cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; width=&quot;30%&quot;&gt;
        '''&lt;tr &gt;
        '''&lt;td class=&quot;ossz01&quot;&gt;NEM SÜRGŐS - FONTOS&lt;/td&gt;
        '''&lt;td class=&quot;ossz02&quot;&gt;[SUMNSF1]&lt;/td&gt;
        '''&lt;td class=&quot;ossz03&quot;&gt;[SUMNSF2]&lt;/td&gt;
        '''&lt;td class=&quot;ossz03&quot;&gt;[SUMNSF3]&lt;/td&gt;
        '''&lt;td class=&quot;ossz03&quot;&gt;[SUMNSF4]&lt;/td&gt;
        '''&lt;/tr&gt;
        '''
        '''&lt;tr &gt;
        '''&lt;td class=&quot;ossz04&quot;&gt;SÜRGŐS - FONTOS&lt;/td&gt;
        '''&lt;td class=&quot;ossz05&quot;&gt;[SUMSF1]&lt;/td&gt;
        '''&lt;td class=&quot;ossz06&quot;&gt;[SUMSF2]&lt;/td&gt;
        '''&lt;td class=&quot;ossz06&quot;&gt;[SUMSF3]&lt;/td&gt;
        '''&lt;td class=&quot;ossz06&quot;&gt;[SUMSF4]&lt;/td&gt;
        '''&lt;/tr&gt;
        '''&lt;tr &gt;
        '''&lt;td class=&quot;auto-style3&quot;&gt;SÜRG [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property ReportSumTable() As String
            Get
                Return ResourceManager.GetString("ReportSumTable", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
