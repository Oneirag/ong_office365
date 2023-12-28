# ONG_OFFICE365

## Summary
Combines [Office365-REST-Python-Client](https://pypi.org/project/Office365-REST-Python-Client/) and https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
to deal with Sharepoint sites with MFA authentication, without asking for password all time.

For some cases (e.g. when you don't have admin rights to create an app or your client_id does not support a certain app), uses selenium to get
tokens from browser and be able to use internal apis.

It is mainly meant for windows (although it also works in macos).

## Prerequisites

In order to access to office365 services, you'll need:
* An email account registered in microsoft
* A tenant name. If you don't know your tenant name, enter in azure portal->Manage Microsoft Entra ID->look for main domain. It will be called `tenant`.onmicrosoft.com
* A client_id/app_id:
  * If you have administrative rights, you can create one in https://go.microsoft.com/fwlink/?linkid=2083908 and copy app_id from there.
  * Otherwise, you can try to use a current client_id from the already registered ones. To do so, get the list of current applications in csv format (using "download" in https://go.microsoft.com/fwlink/?linkid=2083908) and look for a suitable one using the `find_client_ids.py` script
* If you don't have a client_id/app_id, or you cannot create one, then you can use the selenium alternative, that opens a browser and captures tokens from it.

## Configuration

### With client_id
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

# OPTIONAL: Location where sample file will be uploaded
# If not informed, folder from the first file will be used 
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
### Without client_id (using selenium)
Following instructions for installing selenium (https://pypi.org/project/selenium/), after installing selenium you'll need a driver. 
Download chrome driver from https://chromedriver.chromium.org/downloads and place either in your path or in a directory of your choice.
Should it be not installed in PATH, then use the chrome_driver_path in config file to indicate it

In order to use Chrome browser cache, navigate to [chrome://version](chrome://version/) and get the `profile path` from it.
Add it to the config file (`~/.ong_office365.cfg` or `ong_office365.cfg` in current directory):


```ini
[selenium]
profile_path=copy here the chrome profile path. If empty no permanent cache will be used
# Optional
chrome_driver_path=path where chromedriver executable is located
# Optional: pages to block and avoid loading (e.g. put homepage here to avoid opening it)
block_pages=https://www.someserver.com/
```