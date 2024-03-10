# Reminders

Welcome to the 'reminders' project. This started as a tool to help automate sending email based on data in an Excel spreadsheet for my wife. Since I'm not a Visual Basic programmer and don't even have Excel installed on my computer, I reached for Python to help.

## Installation

The program uses some additional Python modules. So, you will need to install. Here's the suggested method for installing the modules:
```sh
python3 -m pip install --user --upgrade -r requirements.txt
```

## Usage

This program includes many fields to generate emails. A configuration file is required to specify the information to generate the email -- it was too much to try and specify all the required parameters as CLI options.

In addition to the spreadsheet and config file option, here are a couple other options that may be useful:
* `-d|--days` - specifies how many days before the item is due. Currently, defaults to 14 days.
* `-p|--person` - used to target at a specific person
* `-i|--interactive` - provides interactive mode where you can skip single users, or just view.

These options may change over time, so consutling the current `--help` is the best way to see what you are doing.

## Configuration File

The configuration file is the means for providing the inputs to generate the emails from an Excel spreadsheet. It has several sections for improved understanding. The sections are:
* **source** - spreadsheet tabs and fields
* **email** - server and address information
* **message** - information included in the message message

More details about each section below.

### Section "source"

This section identifes fields that are important for determinig which action items need a reminder, where to send the reminders, and how to correlate the actions and users. 

Here are the fields:
* **spreadsheet** (optional) - can be used to avoid needing to specify spreadsheet
* **user_id** - Used to correlate users in the actions and users tabs. This field must be present in both tabs
* **tab_users** - Used to identify the spreadsheet tab containing User fields
    * **email_addr** - Column that contains the "target" email address
* **tab_actions** - Used to identify the spreadsheet tab containing Action fields
    * **action_id** - Column that contains the ID. Used as shorthand in some of the parsing
    * **action_due** - Column that contains the due date that is used for identifying actions that need reminders.
    * **action_status** - Column that contains the status. Anything not listed as `Open` in this column will be ignored.


### Section "email"

This section identifies the email parameters.

Here are the fields:
* **server** - SMTP server name
* **port** - SMTP server port
* **from** - email address from which the email will be sent
* **password** (optional) - password will be prompted for (non-echoed) if not provided here
* **cc** - a comma delimted list of email addresses to get copied on each email
* **subject** - subject line for email

### Section "message"

This section defines the body of the emails. Each email body contains a **preamble**, a table using the **columns** from the user actions, and a **close**.

Here are the fields:
* **preamble** - templatized greeting for the message
* **columns** - lists columns in table in the email sent to each user
* **align** - comma delimted field used to change alignment for individual columns.
* **close** - templatized closing for the message

The **preamble** and **close** allow for templating from the User values. The format is HTML for special things (e.g. `<br/>`, `<p/>`). The `{}` denotes a field that will be replaced with something from the User tab. For example, `{First Name}` would get replace with the value from the user's `First Name` column in the user tab.

The **columns** is a comma delimted list of Action columns in the table sent to users. The **align** is used to change from the default alignment (centered); it is a comma delimited list of the form `column-name:alignment` where alignment is `l`, `r`, or `c`.

## Example

There is a short example in `example/`. The example allows you to see how this could be used, and provides a little comedic relief. It is best to run in interactive mode (using `-i|--interactive` flag) so you can see what the table would look like for the provided data.

Here's an example run:
```sh
~/reminders/example> ../reminders.py -c config.ini -i
Barney Rubble: 2
What should be done with Barney Rubble's action?
  skip - skip sending email to this user
  email - send email to barney@slaterockandgravel.com
  show - show the actions (another choice allowed)
  exit - do NOT send anymore emails to anyone
Choose skip, email, show, or exit: show
+---------+------------------+--------------+------------+-------------------------------------------------------+
|    ID   | Action           |     User     |  Due Date  | Notes                                                 |
+---------+------------------+--------------+------------+-------------------------------------------------------+
| Rubble1 | Get married      | Barney/Betty | 1963-06-20 | Left open from a long time ago                        |
|   SG2   | Show up to work  | Fred/ Barney | 1980-03-17 | Need to do this every day, so date is far in the past |
+---------+------------------+--------------+------------+-------------------------------------------------------+
What should be done with Barney Rubble's action?
Choose skip, email, show, or exit: skip
Betty Rubble: 1
What should be done with Betty Rubble's action?
  skip - skip sending email to this user
  email - send email to betty.rubble@hotmail.com
  show - show the actions (another choice allowed)
  exit - do NOT send anymore emails to anyone
Choose skip, email, show, or exit: skip
Fred Flintstone: 1
What should be done with Fred Flintstone's action?
  skip - skip sending email to this user
  email - send email to fred@slaterockandgravel.com
  show - show the actions (another choice allowed)
  exit - do NOT send anymore emails to anyone
Choose skip, email, show, or exit: show
+-----+------------------+--------------+------------+-------------------------------------------------------+
|  ID | Action           |     User     |  Due Date  | Notes                                                 |
+-----+------------------+--------------+------------+-------------------------------------------------------+
| SG2 | Show up to work  | Fred/ Barney | 1980-03-17 | Need to do this every day, so date is far in the past |
+-----+------------------+--------------+------------+-------------------------------------------------------+
What should be done with Fred Flintstone's action?
Choose skip, email, show, or exit: exit
No more users
~/reminders/example> 
```

## TODO
* Write tests
* Create a Docker container

## Contributing

This is an open source project, and I hope you find some utililty in this program (or at least some of the code examples). I would be happy to accept contributions to improve this so others may find it useful, too.
