Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("Savvy Repair for Microsoft Office")> 
<Assembly: AssemblyDescription("Savvy Repair for Microsoft Office, tries four methods for repair or recovery from corruption of Word DOCX, Excel XLSX and PowerPoint PPTX files. DOCX, XSLX and PPTX files are collections of mostly XML sub-files. All four methods first try to repair the zip structure. Additionally in the first method XML corruption is sought and then at first XML error, the sub-files are truncated and repaired. The second method is like the first except its XML file validation for discovering errors is more lax. With the third method, the strict XML validation is returned to and corrupt XML sub-files are truncated and repaired. Additionally with this method, missing XML sub-files are brought in from a blank one of the correct extension. The fourth method uses SilverCoder's DocToText to salvage text or data after which the file is opened as an old style MS Office 97 - 2003 format file including the naked recovered text or data.")> 
<Assembly: AssemblyCompany("S2 Services")> 
<Assembly: AssemblyProduct("Savvy Repair for Microsoft Office, tries four methods for repair or recovery from corruption of Word DOCX, Excel XLSX and PowerPoint PPTX files. DOCX, XSLX and PPTX files are collections of mostly XML sub-files. All four methods first try to repair the zip structure. Additionally in the first method XML corruption is sought and then at first XML error, the sub-files are truncated and repaired. The second method is like the first except its XML file validation for discovering errors is more lax. With the third method, the strict XML validation is returned to and corrupt XML sub-files are truncated and repaired. Additionally with this method, missing XML sub-files are brought in from a blank one of the correct extension. The fourth method uses SilverCoder's DocToText to salvage text or data after which the file is opened as an old style MS Office 97 - 2003 format file including the naked recovered text or data.")> 
<Assembly: AssemblyCopyright("Copyright Paul D Pruitt ©  2013")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("a4a5fb8c-a444-4adf-b5e8-ab15b82fb40e")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.22")> 
<Assembly: AssemblyFileVersion("1.0.0.22")> 

<Assembly: NeutralResourcesLanguageAttribute("en-US")> 