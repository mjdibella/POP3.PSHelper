# POP3.PSHelper
A PowerShell module with helpful cmdlets for accessing the POP3 servers.

Requires the openpop.dll .NET assemply from http://hpop.sourceforge.net/.

To install, create a subdirectory under the $env:PSModulePath directory named POP3.PSHelper and copy the POP3.PsHelper.psm1 and openpop.dll file into that directory. Then create a shortcut to open a POP3 Powershell window with the command line:

powershell.exe -noExit -Command "& {Import-module POP3.PSHelper.psm1}"
