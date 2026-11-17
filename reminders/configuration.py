from configparser import ConfigParser
from typing import List


###############################################################################
###############################################################################
CSECT_EMAIL = 'email'
CSECT_MSG = 'message'
CSECT_SOURCE = 'source'


class SafeConfigParser(ConfigParser):
    """
    Class with a get that does NOT throw when item does NOT exist
    """
    def get(self, section: str, option: str, raw: bool = True, **kwargs) -> str:
        if not self.has_option(section, option):
            return ''
        return super().get(section, option, raw=raw, **kwargs)


class Configuration:
    """
    Spreadsheet, email, and message settings used to build reminder emails.
    """
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
