# Block-Office-Macros-on-16.0
A PowerShell script to automate the creation of a GPO using Office 2016 Administrative Templates that will block the execution of all Office macros when downloaded from the internet.

This script begins with the verification of the presence of a central store for policy definitions. It can only be run on the server that is the holder of the domain's FSMO roles, so if you are running with two or more domain controllers, it will self-exit if it determines itself to be run under a non-FSMO server. 

Once it creates the central store, it will check if you currently have the 2020 Office Administrative Templates (.ADMX), and if it doesn't, it will download them. 

If you are missing the newest admin templates, the script will download them, extract them to the central store. If you are using a domain with a language/locale other than EN-US, this script needs to be modified at line 71 to include multiple locales or to replace the en-us locale. 

This script will then use the new ADMX templates to create the Group Policy Objects to restrict the use of Macros from downloaded Office files (filetype .doc, .docs, .ppt, .pptx, .xls, .xlsx). In testing, before this GPO was implemented, running a Macro from a downloaded Word document was possible, and only required the ignoring of a few warnings and the suspension of a hold that Defender put onto the downloaded document in Chrome. It is likely that if the file was sourced from the DC or an email, it would not catch the malicious download, but it would still give ample warning. Without the GPO, the script will try to execute and with minimal intervention, it does. 

However, when this GPO is enabled, any execution of any macro is explicitly prevented, and even prompts with a "You do not have the required permissions to do that" pop-up. 

The script will then cleanup temp files.
