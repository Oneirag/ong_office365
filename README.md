ONG_OFFICE365
=============

# Summary
Combines [Office365-REST-Python-Client](https://pypi.org/project/Office365-REST-Python-Client/) and https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
to deal with Sharepoint sites with MFA authentication, without asking for password all time.

For some cases (e.g. when you don't have admin rights to create an app or your client_id does not support a certain app), uses selenium to get
tokens from browser and be able to use internal apis.

It is mainly meant for windows (although it also works in macos).

# Prerequisites

In order to access to office365 services, you'll need:
* An email account registered in microsoft

Then you have two alternatives:
* If you have a client_id/app_id:
  * A tenant name. If you don't know your tenant name, enter in azure portal->Manage Microsoft Entra ID->look for main domain. It will be called `tenant`.onmicrosoft.com
  * If you have administrative rights, you can create one in https://go.microsoft.com/fwlink/?linkid=2083908 and copy app_id from there.
  * Otherwise, you can try to use a current client_id from the already registered ones. To do so, get the list of current applications in csv format (using "download" in https://go.microsoft.com/fwlink/?linkid=2083908) and look for a suitable one using the `find_client_ids.py` script
* **If you don't have a client_id/app_id, or you cannot create one**: you can use the selenium alternative, that opens a browser and captures tokens from it.

# Configuration

## With client_id
In order for the software to work, a .yaml file should be created with the name `~/.config/ongpi/ong_office365.yaml` or `ong_office365.yaml` in current directory.

Example content of the file:

```yaml
ong_office365:
  # Leave empty (null) to get email from outlook in windows
  email: contoso@your-tenant-name.com
  # Use tenant from email. Typically is your email domain address
  # Use common for personal Microsoft accounts and work/school accounts from Azure Active Directory
  # Use organizations for work/school accounts from Azure Active Directory only
  # Use consumers for personal Microsoft accounts (MSA) only 
  tenant: your-tenant-name
  # Example of a sharepoint personal site (onedrive)
  # sharepoint = https://${tenant}-my.sharepoint.com/personal/{your_parsed_email}
  # client_id should come from https://go.microsoft.com/fwlink/?linkid=2083908
  # This is a sample value from a google search 
  client_id: 6731de76-14a6-49ae-97bc-6eba6914391e
  # Optional value. Seconds to keep interactive window open. Defaults to 20
  timeout: 10
  sharepoint:
    site_url: https://${tenant}.sharepoint.com/sites/{site}
    client_id:
            - client_id1
            - client_id2
  
    # OPTIONAL: Location where sample file will be uploaded
    # If not informed, folder from the first file will be used 
    dest_url: Shared Documents/{folder1}/{folder2}
  
  onedrive:
    # Overrides default section for Sharepoint class
    # Defaults to first one, list is used in tests
    client_id:
            - client_id1
            - client_id2
    # OPTIONALS: just for tests
    # Urls of files in the sharepoint that will be test to be downloaded
    relative_urls:
            - /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename1
            - /sites/{site-name}/Shared Documents/{folder1}/{folder2/filename2
    # Location where sample file will be uploaded
    dest_url: Shared Documents/{folder1}/{folder2}
```
## Without client_id (using selenium)
Following instructions for installing selenium (https://pypi.org/project/selenium/), after installing selenium you'll need a driver. 
Download chrome driver from https://chromedriver.chromium.org/downloads and place either in your path or in a directory of your choice.
Should it be not installed in PATH, then use the chrome_driver_path in config file to indicate it

In order to use Chrome browser cache, navigate to [chrome://version](chrome://version/) and get the `profile path` from it.
Add it to the yaml config file (`~/.config/ongpi/ong_office365.yaml` or `ong_office365.yaml` in current directory):


```yaml
ong_office365:
  sharepoint:
    # Needed to access to sharepoint. Typically, in the form of https://${tenant}.sharepoint.com/sites/{site}
    site_url: whateversite
  selenium:
    # Required. If path contains {user} will be replaced with current username. If you want to use chrome seetings
    # Navigate to chrome://version and copy-paste profile path from it
    profile_path: copy here the chrome profile path. If empty no permanent cache will be used
    # Optional
    chrome_driver_path: path where chromedriver executable is located
    # Optional: pages to block and avoid loading (e.g. put homepage here to avoid opening it)
    block_pages: https://www.someserver.com/
```

# Use of ms forms
Access ms forms can only be performed using selenium. See sample config file [here](#without-clientid-using-selenium)

## Sample code:
```python
from ong_office365.ong_forms import Forms

forms = Forms()

##################
# list all forms
##################
all_forms = forms.get_forms()
print(all_forms)
###################################
# Get responses for a certain form
###################################
# form_id = all_forms[-1]['id']
title_to_search_for = "One title to search for"
my_form = forms.get_form_by(title=title_to_search_for)
if my_form:
    form_id = my_form[-1]['id']
    df = forms.get_pandas_result(form_id)
    # each answer is a row, questions are in columns
    print(df)
else:
    print(f"Form '{title_to_search_for}' not found")

#############################################
# Create a simple form, with sample questions
#############################################
from ong_office365.forms_objects.questions import QuestionText, QuestionChoice
import datetime
new_form = forms.create_form(title="Sample form: " + datetime.datetime.now().isoformat())
new_form_id = new_form['id']

# Add multiple line question. Has no subtitle
question_text = QuestionText(title=f"Long text question", multiline=True)
q_text = forms.create_question(new_form_id, question_text)
# Add a choice question with radio buttons so only one answer can be selected
question_option = QuestionChoice(title=f"Choice question",
                                 choices=["One", "Two"], subtitle="Select one")
q_option = forms.create_question(new_form_id, question_option)
# Add a one line question
question_text = QuestionText(title=f"Short text question", multiline=False,
                             subtitle="Just one line")
q_text = forms.create_question(new_form_id, question_text)
# Add a multiple choice that has check buttons to select multiple. It is mandatory (required=True)
question_option = QuestionChoice(title=f"Mandatory multiple choice question",
                                 choices=["One", "Two", "three"], subtitle="Select multiple",
                                 allow_other_answer=True, required=True)
q_option = forms.create_question(new_form_id, question_option)

forms.trash_form(new_form_id)
# forms.delete_form(new_form_id)

##########################################################
# Create a simple form, with sample questions and sections
##########################################################
from ong_office365.forms_objects.questions import Section
new_form = forms.create_form(title="Sample form with sections: " + datetime.datetime.now().isoformat())
new_form_id = new_form['id']

#############################################
# following questions appear under section 1
#############################################
section1 = Section(title="First section", subtitle="Sample section")
q_section = forms.create_section(new_form_id, section1)
# Add multiple line question. Has no subtitle
question_text = QuestionText(title=f"Long text question", multiline=True)
q_text = forms.create_question(new_form_id, question_text)
# Add a choice question with radio buttons so only one answer can be selected
question_option = QuestionChoice(title=f"Choice question",
                                 choices=["One", "Two"], subtitle="Select one")
q_option = forms.create_question(new_form_id, question_option)

#############################################
# following questions appear under section 2
#############################################
section2 = Section(title="Second section", subtitle="Another sample section")
q_section = forms.create_section(new_form_id, section2)
# Add a one line question
question_text = QuestionText(title=f"Short text question", multiline=False,
                             subtitle="Just one line")
q_text = forms.create_question(new_form_id, question_text)
# Add a multiple choice that has check buttons to select multiple. It is mandatory (required=True)
question_option = QuestionChoice(title=f"Mandatory multiple choice question",
                                 choices=["One", "Two", "three"], subtitle="Select multiple",
                                 allow_other_answer=True, required=True)
q_option = forms.create_question(new_form_id, question_option)

forms.trash_form(new_form_id)
# forms.delete_form(new_form_id)

##########################################################
# Complex example with sections and a menu with branches
##########################################################
# Creates a first section "Menu". Then creates more sections, that move to the end after each one
# Finally creates a ChoiceQuestion in the first section that sends to the corresponding section
new_form = forms.create_form("Complex form with sections and branches: " + datetime.datetime.now().isoformat(),
                             description="Subtitle goes here",
                             )
new_form_id = new_form['id']
# Create a first section, where menu will be hosted
section = Section(title="Main menu", subtitle="Choose section")
menu_section = forms.create_section(new_form_id, section)
# Create sections
sections = []
for section_id in range(3):
    # Creates a new section that jumps directly to the end after filling it
    section = Section(title=f"Section {section_id}", subtitle=f"Sample section {section_id}",
                      to_the_end=True)
    q_section = forms.create_section(new_form_id, section)
    sections.append(q_section)
    # Add some questions to the section
    question_text = QuestionText(title=f"Long text question of section {section_id}", multiline=True)
    q_text = forms.create_question(new_form_id, question_text)
    question_option = QuestionChoice(title=f"Choice question of section {section_id}",
                                     choices=["One", "Two"], subtitle="Select one")
    q_option = forms.create_question(new_form_id, question_option)
    question_text = QuestionText(title=f"Short text question for section {section_id}", multiline=False,
                                 subtitle="Just one line")
    q_text = forms.create_question(new_form_id, question_text)
    question_option = QuestionChoice(title=f"Multiple choice question of section {section_id}",
                                     choices=["One", "Two", "three"], subtitle="Select multiple",
                                     allow_other_answer=True)
    q_option = forms.create_question(new_form_id, question_option)

# Now, create a question menu that sends to each of the branches (sections)
# Note that an order between section1 and section2 order must be informed. Default value
# would add question to the last created section
menu = QuestionChoice(title="", choices=sections, order=menu_section['order'] + 1)
q_menu = forms.create_question(new_form_id, menu)

forms.trash_form(new_form_id)
```



