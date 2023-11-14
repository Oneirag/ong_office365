# ONG_OFFICE365

## Summary
Combines [Office365-REST-Python-Client](https://pypi.org/project/Office365-REST-Python-Client/) and https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
to deal with Sharepoint sites with MFA authentication, without asking for password all time.

It is mainly meant for windows (although it also works in macos)

## Configuration

In order for the software to work, a .cfg file should be created with the name `~/.ong_office365.cfg` or `ong_office365.cfg` in current directory.

Example content of the file:

```ini
[DEFAULT]
# Leave empty to get email from outlook in windows
email = contoso@your-tenant-name.com
# Use tenant from email. Typically is your email domain address
tenant = your-tenant-name
# Example of a sharepoint site
sharepoint = https://${tenant}.sharepoint.com/sites/{site}
# Example of a sharepoint personal site (onedrive)
# sharepoint = https://${tenant}-my.sharepoint.com/personal/{your_parsed_email}
# client_id should come from https://go.microsoft.com/fwlink/?linkid=2083908
# This is a sample value from a google search 
client_id = 6731de76-14a6-49ae-97bc-6eba6914391e
```
In order to run tests, an additional section "tests" would be required
```ini
[tests]
client_ids =
        client_id1
        client_id2
# Urls of files in the sharepoint that will be test to be downloaded
relative_urls =
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename1
        /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename2
# Location where sample file will be uploaded
dest_url = Shared Documents/{folder1}/{folder2}

```
