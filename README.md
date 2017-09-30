# Powershell Scripts
This repo contains Powershell scripts used for several functions.

At the moment it consists mainly of:
- a Login script, which includes functions to facilitate logging in to Exchange Online, MSOnline and Sharepoint Online services;
- a GeneralFunctions module, which contain mostly functions which interact with these services.

It is important to keep them in a single folder when using them, as many functions rely on the path created by the `$global:GeneralRoot` variable for their logic.
