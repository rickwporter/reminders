#!/usr/bin/env python3
import argparse
import datetime
import prettytable
import re
import sys

from configparser import ConfigParser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from enum import Enum
from getpass import getpass
from pandas import ExcelFile
from pathlib import Path
from smtplib import SMTP
from typing import Any
from typing import Dict
from typing import List
from typing import Optional
from typing import Set
from typing import Tuple


###############################################################################
###############################################################################
CSECT_EMAIL = 'email'
CSECT_MSG = 'message'
CSECT_SOURCE = 'source'

ROW_HEADER = '_row'

FIELD_RE = re.compile(r'\{.*\}')
NL = '\n\t'


#####################
# constants for program
class Format(str, Enum):
    HTML = 'html'
    JSON = 'json'
    CSV = 'csv'
    TEXT = 'text'


# what to do in interactive mode
class What(str, Enum):
    SKIP = 'skip'
    EMAIL = 'email'
    SHOW = 'show'
    EXIT = 'exit'


class MissingUser(Exception):
    """
    Simple wrapper to format a missing user correctly.
    """
    def __init__(self, username: str, record: str, row: str):
        message = f"Missing '{username}' on {record}, row={row}"
        super().__init__(message)


class AmbiguousUser(Exception):
    """
    Simple wrapper to format errors
    """
    def __init__(self, username: str, matches: List[str]):
        message = f"Ambiguous result for '{username}' matches: {', '.join(matches)}"
        super().__init__(message)


class SafeConfigParser(ConfigParser):
    """
    Class with a get that does NOT throw when item does NOT exist
    """
    def get(self, section: str, option: str, raw: bool = True, **kwargs) -> str:
        if not self.has_option(section, option):
            return ''
        return super().get(section, option, raw=raw, **kwargs)


###############################################################################
# Function definitions
#    These should not be modified
###############################################################################
class Reminders:
    def __init__(self):
        self.spreadsheet = None
        self.tab_user = None
        self.tab_action = None
        self.hdr_user = None
        self.hdr_email = None
        self.hdr_id = None
        self.hdr_due = None
        self.hdr_status = None
        self.mail_server = None
        self.mail_port = None
        self.mail_from = None
        self.mail_password = None
        self.mail_subject = None
        self.mail_cc = None
        self.msg_preamble = None
        self.msg_table_headers = None
        self.msg_table_align = None
        self.msg_close = None

    def parse_args(self, *args) -> argparse.Namespace:
        """
        Parses the provided arguments into a structure provided by argparse.
        """
        parser = argparse.ArgumentParser(
            description="Parse the specified Excel file looking for action items to remind."
        )
        parser.add_argument(
            '-s',
            '--spreadsheet',
            dest="spreadsheet",
            type=str,
            help="Excel spreadsheet to query.",
        )
        parser.add_argument(
            '-c',
            '--config',
            dest='config',
            type=str,
            help="Configuration file"
        )
        parser.add_argument(
            "-p",
            "--person",
            type=str,
            help="Specify person for whom to generate reminders",
        )
        parser.add_argument(
            "-d",
            "--days",
            type=int,
            default=14,
            help="Number of days ahead of deadline to warn (default : %(default)s).",
        )
        parser.add_argument(
            "-i",
            "--interactive",
            action="store_true",
            help="Interactive mode allows viewing per user data before sending."
        )
        return parser.parse_args(*args)

    def print(self, *args, **kwargs) -> None:
        """
        This is used for easier mocking/capture in unittests
        """
        print(*args, **kwargs)

    def parse_config(self, filename: str) -> None:
        config = SafeConfigParser()
        config.read(filename)
        self.update_config(config)

    def update_config(self, config: SafeConfigParser) -> None:
        self.spreadsheet = config.get(CSECT_SOURCE, 'spreadsheet') or self.spreadsheet
        self.tab_user = config.get(CSECT_SOURCE, 'tab_users') or self.tab_user
        self.tab_action = config.get(CSECT_SOURCE, 'tab_actions') or self.tab_action
        self.hdr_user = config.get(CSECT_SOURCE, 'user_id') or self.hdr_user
        self.hdr_email = config.get(CSECT_SOURCE, 'email_addr') or self.hdr_email
        self.hdr_id = config.get(CSECT_SOURCE, 'action_id') or self.hdr_id
        self.hdr_due = config.get(CSECT_SOURCE, 'action_due') or self.hdr_due
        self.hdr_status = config.get(CSECT_SOURCE, 'action_status') or self.hdr_status

        self.mail_server = config.get(CSECT_EMAIL, 'server') or self.mail_server
        self.mail_port = config.get(CSECT_EMAIL, 'port') or self.mail_port
        self.mail_password = config.get(CSECT_EMAIL, 'password') or self.mail_password
        self.mail_from = config.get(CSECT_EMAIL, 'from') or self.mail_from
        self.mail_subject = config.get(CSECT_EMAIL, 'subject') or self.mail_subject
        self.mail_cc = [_.strip() for _ in config.get(CSECT_EMAIL, 'cc').split(',') if _.strip()] or self.mail_cc

        self.msg_preamble = config.get(CSECT_MSG, 'preamble') or self.msg_preamble
        self.msg_close = config.get(CSECT_MSG, 'close') or self.msg_close
        self.msg_table_headers = [
            _.strip() for _ in config.get(CSECT_MSG, 'columns').split(',') if _.strip()
        ] or self.msg_table_headers
        avalues = [_.strip() for _ in config.get(CSECT_MSG, 'align').split(',') if _.strip()]
        if avalues:
            self.msg_table_align = {}
            for a in avalues:
                parts = [_.strip() for _ in a.split(':') if _.strip()]
                if len(parts) != 2:
                    raise ValueError(f"The {CSECT_MSG}/align option must be of the form 'header: align-value'")
                name = parts[0]
                align = parts[1]
                self.msg_table_align.update({name: align})
        return

    def check_config(self) -> List[str]:
        """
        Insures all the "required" fields have been filled in.

        Returns a list with any errors.
        """
        errors = []
        if not self.tab_action:
            errors += ["Missing spreadsheet tab name for actions"]
        if not self.tab_user:
            errors += ["Missing spreadsheet tab name for users"]
        if not self.hdr_user:
            errors += ["Missing user/action user-id field"]
        if not self.hdr_email:
            errors += ["Missing user email field"]
        if not self.hdr_id:
            errors += ["Missing action identifier field"]
        if not self.hdr_due:
            errors += ["Missing action due date field"]
        if not self.hdr_status:
            errors += ["Missing action status field"]
        if not self.mail_server or not self.mail_port:
            errors += ["Missing mail server or port"]
        if not self.mail_from:
            errors += ["Missing mail from address"]
        if not self.mail_subject:
            errors += ["Missing mail subject"]
        if not self.msg_table_headers:
            errors += ["Missing message table headers"]
        # TODO: check that preamble and closing exist?
        return errors

    def sheet_to_dict(self, filename: str, sheetname: str) -> List[Dict]:
        """
        Open the Excel file and parse the items in the specified sheet to a list of
        dictionaries, where the dictionary property names align with the sheet header
        column strings.
        """
        with ExcelFile(filename) as xls:
            sheet = xls.parse(sheetname)
            values = sheet.to_dict()
            headers = values.keys()
            num_values = max([len(_.keys()) for _ in values.values()])

            result = []
            for row in range(num_values):
                item = {k: values.get(k, {}).get(row, '') for k in headers}
                item[ROW_HEADER] = f"{sheetname}:{row + 1}"
                result.append(item)

            return result

    def find_user(self, users: List[Dict], username: str) -> Dict:
        """
        Search the whole list of users to find a user who has something in a field that matches (case-insensitive)
        """
        matches = []
        searchname = username.lower().strip()
        for user in users:
            for v in user.values():
                if not isinstance(v, str):
                    continue

                if searchname in v.lower():
                    # if we find a match, break out of the inner loop
                    matches.append(user)
                    break

        if not matches:
            return None

        if len(matches) > 1:
            raise AmbiguousUser(username, [_.get(ROW_HEADER) for _ in matches])

        return matches[0]

    def correlate(self, users: List[Dict], actions: List[Dict]) -> List[Tuple[Dict, List[Dict]]]:
        """
        Organizes the actions by user. It returns a list of tuples(user, list(actions))

        To avoid issues with hashing a dictionary, a Tuple is used. However, to collect
        the items for a given tuple, a dictionary is used internally and converted to a
        list of tuples as the last step.
        """
        uname_actions = {}
        for action in actions:
            for username in action.get(self.hdr_user).split('/'):
                user = self.find_user(users, username)
                if not user:
                    raise MissingUser(username, action.get(self.hdr_id), action.get(ROW_HEADER))
                uname = user.get(self.hdr_user)
                uact = uname_actions.get(uname, [])
                uact.append(action)
                uname_actions.update({uname: uact})

        # convert the internal uname_actions dictionary into a list of tuple(user, list(actions))
        return [(self.find_user(users, uname), uact) for uname, uact in uname_actions.items()]

    def _format(self, value: Any) -> str:
        """
        Creates strings out of each cell/value
        """
        # now, try to the cell/data value
        if isinstance(value, datetime.datetime):
            value = str(value.date())
        return str(value)

    def _get_time(self, value: datetime.datetime) -> Optional[datetime.date]:
        """
        Pulls the date out of the provided value
        """
        if isinstance(value, datetime.datetime):
            return value.date()
        return None

    def substitute(self, template: str, user: Dict, days: int) -> str:
        """
        Substitutes the {}'s in the template with either the days, or a value
        from the user.

        For example, a '{First Name}' would get the value of 'user.get("First Name")'.
        If the field name does not exist in user (and not equal to 'days'), a ValueError
        is raised (instead of sending email with an unknown/missing value)
        """
        def new_value(group: str) -> str:
            # the group includes the braces, so trim to avoid the braces
            v = group[1:-1]
            if v == 'days':
                return str(days)
            if v in user:
                return user.get(v)
            raise ValueError(f"Missing user field='{v}'")

        return FIELD_RE.sub(lambda m: new_value(m.group()), template)

    def _create_table(self, actions: List[Dict], fmt: Format) -> str:
        """
        Creates a formatted table out of the list of actions.
        """
        table = prettytable.PrettyTable()
        table.field_names = self.msg_table_headers
        # horizontal alignment can be overridden
        for h, v in self.msg_table_align.items():
            table.align[h] = v

        for item in actions:
            # remove the time, and turn the datetime value into a string
            values = [self._format(item.get(h)) for h in self.msg_table_headers]
            table.add_row(values)

        if fmt == Format.HTML:
            return table.get_html_string(
                header=True,
                border=True,
                hrules=prettytable.ALL,
                vrules=prettytable.ALL,
                format=True,
            )
        if fmt == Format.CSV:
            return table.get_csv_string()
        if fmt == Format.JSON:
            return table.get_json_string()
        return table.get_string()

    def get_email_server(self) -> SMTP:
        """
        Simple "utility" to log into a mail server
        """
        # log into the email server once
        password = self.mail_password or getpass(f"{self.mail_from} email password:")
        server = SMTP(self.mail_server, self.mail_port)
        server.starttls()
        server.login(self.mail_from, password)
        return server

    def send_email_via_server(self, server: SMTP, user: Dict, actions: List[Dict], days: int) -> None:
        """
        Formats the message to the user, and sends it using the server.
        """
        to_user = user.get(self.hdr_email)
        intro = self.substitute(self.msg_preamble, user, days)
        closing = self.substitute(self.msg_close, user, days)
        table = self._create_table(actions, Format.HTML)
        body = f"{intro}{table}{closing}"
        to = [to_user]
        if isinstance(self.mail_cc, str):
            to.append(self.mail_cc)
        elif isinstance(self.mail_cc, list):
            to.extend(self.mail_cc)

        message = MIMEMultipart("alternative")
        message['Subject'] = self.mail_subject
        message['From'] = self.mail_from
        message['To'] = to_user
        message['CC'] = self.mail_cc
        message.attach(MIMEText(body, "html"))
        server.sendmail(self.mail_from, to, message.as_string())
        return

    def send_all_emails(self, user_actions: Dict, days: int) -> None:
        """
        Send emails to all the users in the list with actions.
        """
        self.print(f"Sending emails about items due in the next {days} days:")
        for (user, actions) in user_actions:
            self.print(f"    {user.get(self.hdr_user)}: {len(actions)}")

        # each user gets a tailored email
        server = self.get_email_server()
        for (user, actions) in user_actions:
            self.send_email_via_server(server, user, actions, days)

        # close things out nicely
        server.quit()
        return

    def prompt_for_what(self, user: Dict, actions: List[Dict]) -> What:
        """
        Provides menu for what to do with a particular user/actions.
        """
        self.print(f"{user.get(self.hdr_user)}: {len(actions)}")
        what = None
        while what not in (What.SKIP, What.EMAIL, What.EXIT):
            self.print(f"What should be done with {user.get(self.hdr_user)}'s action?")
            if not what:
                self.print(f"  {What.SKIP} - skip sending email to this user")
                self.print(f"  {What.EMAIL} - send email to {user.get(self.hdr_email)}")
                self.print(f"  {What.SHOW} - show the actions (another choice allowed)")
                self.print(f"  {What.EXIT} - do NOT send anymore emails to anyone")
            what = input(f"Choose {What.SKIP}, {What.EMAIL}, {What.SHOW}, or {What.EXIT}: ").lower()
            if what == What.SHOW:
                self.print(f"{self._create_table(actions, Format.TEXT)}")
        return what

    def interactive_send_email(self, user_actions: List[Tuple[Dict, List[Dict]]], days: int) -> None:
        """
        Walks the user/actions, prompts what to do for each user, and takes the
        appropriate action.
        """
        server = None
        # each user gets a tailored email
        for (user, actions) in user_actions:
            what = self.prompt_for_what(user, actions)
            if what == What.EXIT:
                break
            if what == What.EMAIL:
                server = server or self.get_email_server()
                self.send_email_via_server(server, user, actions, days)

        self.print('No more users')

        if server:
            server.quit()
        return

    def get_fields(self, string: str) -> Set[str]:
        return set([_.replace('{', '').replace('}', '') for _ in FIELD_RE.findall(string)])

    def valid_string(self, value: Any) -> bool:
        if isinstance(value, str):
            return bool(value)
        return False

    def valid_date(self, value: Any) -> bool:
        if isinstance(value, datetime.date):
            return True
        return False
        
    def validate_users(self, users: List[Dict]) -> List[str]:
        """
        Verifies all users have all "required" fields
        """
        fields = self.get_fields(self.msg_preamble) | self.get_fields(self.msg_close)
        fields.discard('days')  # not part of user record
        fields.discard(self.hdr_email)  # already check for missing email

        errors = []
        for user in users:
            reasons = []
            if not self.valid_string(user.get(self.hdr_email)):
                reasons.append('missing email')
            missing = fields - user.keys()
            if missing:
                reasons.append(f"missing email fields {'/'.join(sorted(missing))}")
            if reasons:
                errors.append(f"{user.get(self.hdr_user)} ({user.get(ROW_HEADER)}) error(s): {', '.join(reasons)}")

        return errors

    def validate_actions(self, actions: List[Dict]) -> List[str]:
        fields = set(self.msg_table_headers)
        # remove the important fields we already check for (to avoid redundant errors)
        fields.discard(self.hdr_user)
        fields.discard(self.hdr_due)

        errors = []
        for action in actions:
            reasons = []
            if not self.valid_string(action.get(self.hdr_user)):
                reasons.append('missing assignment')
            if not self.valid_date(action.get(self.hdr_due)):
                reasons.append('missing due date')
            missing = fields - action.keys()
            if missing:
                reasons.append(f"missing table fields {'/'.join(sorted(missing))}")
            if reasons:
                errors.append(f"{action.get(self.hdr_id)} ({action.get(ROW_HEADER)}) error(s): {', '.join(reasons)}")

        return errors

    def run(self, *sysargs) -> int:
        """
        This is the "main" function that parses args, collects data, and prints it.
        """
        args = self.parse_args(*sysargs)
        days = args.days

        if args.config:
            config = Path(args.config)
            if not config.is_file():
                self.print(f"{args.config} is not a file")
                return 4
            self.parse_config(args.config)

        errors = self.check_config()
        if errors:
            self.print(f"Configuration errors:{NL}{NL.join(errors)}")
            return 5

        spreadsheet = Path(args.spreadsheet or self.spreadsheet)
        if not spreadsheet.is_file():
            self.print(f"{spreadsheet} is not a file!")
            return 1

        users = self.sheet_to_dict(str(spreadsheet), self.tab_user)
        errors = self.validate_users(users)
        if errors:
            self.print(f"Invalid users: {NL}{NL.join(errors)}")
            return 2

        actions = self.sheet_to_dict(str(spreadsheet), self.tab_action)
        errors = self.validate_actions(actions)
        if errors:
            self.print(f"Invalid actions: {NL}{NL.join(errors)}")
            return 3

        # start by filtering out non-Open items... may contain users no longer in the system
        actions = [_ for _ in actions if _.get(self.hdr_status) == 'Open']

        # remove actions not due, before correlating with by users
        before = datetime.datetime.now() + datetime.timedelta(days=days)
        actions = [_ for _ in actions if _.get(self.hdr_due) < before]

        # sort by time here, so it is displayed in order
        actions = sorted(actions, key=lambda i: self._get_time(i[self.hdr_due]))

        user_actions = self.correlate(users, actions)

        # reduce the list to focus only on the specified people
        if args.person:
            user = self.find_user(users, args.person)
            user_actions = [(u, a) for (u, a) in user_actions if u == user]

        # check after filtering
        if not user_actions:
            filter_info = f" for {args.person}" if args.person else ""
            self.print(
                f"No open user actions found in {str(spreadsheet)}"
                f"{filter_info} in the next {days} days"
            )
            return 0

        if not args.interactive:
            self.send_all_emails(user_actions, days)
        else:
            self.interactive_send_email(user_actions, days)
        return 0


if __name__ == "__main__":
    reminders = Reminders()
    result = reminders.run(sys.argv[1:])
    sys.exit(result)
