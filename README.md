# ONG_OFFICE365

## Summary
Combines [Office365-REST-Python-Client](https://pypi.org/project/Office365-REST-Python-Client/) and https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
to deal with Sharepoint sites with MFA authentication, without asking for password all time.

It is mainly meant for windows (although it also works in macos).

## Prerequisites

In order to access to office365 services, you'll need:
* An email account registered in microsoft
* A tenant name. If you don't know your tenant name, enter in azure portal->Manage Microsoft Entra ID->look for main domain. It will be called `tenant`.onmicrosoft.com
* A client_id/app_id:
  * If you have administrative rights, you can create one in https://go.microsoft.com/fwlink/?linkid=2083908 and copy app_id from there.
  * Otherwise, you can try to use a current client_id from the already registered ones. To do so, get the list of current applications in csv format (using "download" in https://go.microsoft.com/fwlink/?linkid=2083908) and look for a suitable one using the `find_client_ids.py` script


## Configuration

In order for the software to work, a .cfg file should be created with the name `~/.ong_office365.cfg` or `ong_office365.cfg` in current directory.

Example content of the file:

```ini
[DEFAULT]
# Leave empty to get email from outlook in windows
email = contoso@your-tenant-name.com
# Use tenant from email. Typically is your email domain address
# Use common for personal Microsoft accounts and work/school accounts from Azure Active Directory
# Use organizations for work/school accounts from Azure Active Directory only
# Use consumers for personal Microsoft accounts (MSA) only 
tenant = your-tenant-name
# Example of a sharepoint site
site_url = https://${tenant}.sharepoint.com/sites/{site}
# Example of a sharepoint personal site (onedrive)
# sharepoint = https://${tenant}-my.sharepoint.com/personal/{your_parsed_email}
# client_id should come from https://go.microsoft.com/fwlink/?linkid=2083908
# This is a sample value from a google search 
client_id = 6731de76-14a6-49ae-97bc-6eba6914391e
# Optional value. Seconds to keep interactive window open. Defaults to 20
timeout = 10
[sharepoint]
# Overrides default section for Sharepoint class
site_url = https://${tenant}.sharepoint.com/sites/{site}
client_id =
        client_id1
        client_id2

# OPTIONALS: Just for tests
# Urls of files in the sharepoint that will be test to be downloaded
relative_urls =
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename1
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename2
# Location where sample file will be uploaded
dest_url = Shared Documents/{folder1}/{folder2}

[onedrive]
# Overrides default section for Sharepoint class
# It must be blank for onedrive
site_url =
# Defaults to first one, list is used in tests
client_id =
        client_id1
        client_id2
# OPTIONALS: just for tests
# Urls of files in the sharepoint that will be test to be downloaded
relative_urls =
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename1
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename2
# Location where sample file will be uploaded
dest_url = Shared Documents/{folder1}/{folder2}
```
