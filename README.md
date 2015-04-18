### Manifest Resource Editor v1.01 ###

The purpose of this tool is to prepare an existing application that consumes  
COM objects from DLLs located inside of or below the folder from where it is  
run for working without registration of involved DLLs, by using isolated COM.  
  
Usage:  
  
manfred <target> ... [ /once ] [ /ini ... ] [ /files ... ] [ /minus ... ]  
  
<target>  may be followed by a list of subfolders to search  
/once     causes update of manifest to occur only when no file tags exist yet  
/never    causes update of manifest to occur never; useful with /rgs option  
/ini      specifies an ini file from which to merge content into the manifest  
/rgs      specifies an rgs file to write results to for bulk registration  
/files    specifies file inclusion patterns; may occur repeatedly  
/minus    specifies file exclusion patterns; may occur only once  
  
Use manfred.exe with 32-bit applications, and manphred.exe with 64-bit ones.
