import unittest

from reminders import CSECT_EMAIL
from reminders import CSECT_MSG
from reminders import Configuration
from reminders import SafeConfigParser


class TestConfig(unittest.TestCase):
    def test_config_parser_empty(self):
        uut = SafeConfigParser()
        uut.read_string("")
        self.assertEqual('', uut.get("foo", "bar"))

    def test_config_parser_basic(self):
        uut = SafeConfigParser()
        uut.read_string("[sna]\nfoo = bar")
        self.assertEqual('bar', uut.get('sna', 'foo'))
        self.assertEqual('', uut.get('sna', 'bar'))  # section exists
        self.assertEqual('', uut.get('blah', 'bar'))  # section does NOT exists

    def test_config_parser_hash(self):
        uut = SafeConfigParser()
        uut.read_string("[sna]\nfoo = #bar")

        # we get back the hash value
        self.assertEqual('#bar', uut.get('sna', 'foo'))

    def test_config_example(self):
        uut = Configuration()
        uut.parse_config('example/config.ini')
        self.assertEqual('bedrock.xlsx', uut.spreadsheet)
        self.assertEqual('Users', uut.tab_user)
        self.assertEqual('Actions', uut.tab_action)
        self.assertEqual('User', uut.hdr_user)
        self.assertEqual('Email', uut.hdr_email)
        self.assertEqual('ID', uut.hdr_id)
        self.assertEqual('Due Date', uut.hdr_due)
        self.assertEqual('Status', uut.hdr_status)

        self.assertEqual('smtp.office365.com', uut.mail_server)
        self.assertEqual('587', uut.mail_port)
        self.assertEqual(None, uut.mail_password)
        self.assertEqual('slate@slaterockandgravel.com', uut.mail_from)
        self.assertEqual('Do It! Now!!!', uut.mail_subject)
        self.assertEqual(None, uut.mail_cc)

        self.assertIn('The following actions assigned to you', uut.msg_preamble)
        self.assertEqual('<p/><p>Thank you,</p><p>Mr. Slate<br/>Slate Rock and Gravel</p>', uut.msg_close)
        self.assertEqual(['ID', 'Action', 'User', 'Due Date', 'Notes'], uut.msg_table_headers)
        self.assertEqual({'Action': 'l', 'Notes': 'l'}, uut.msg_table_align)

        self.assertEqual([], uut.check_config())

    def test_config_empty(self):
        uut = Configuration()
        errors = uut.check_config()
        self.assertIn('Missing spreadsheet tab name for actions', errors)
        self.assertIn('Missing spreadsheet tab name for users', errors)
        self.assertIn('Missing user/action user-id field', errors)
        self.assertIn('Missing user email field', errors)
        self.assertIn('Missing action identifier field', errors)
        self.assertIn('Missing action due date field', errors)
        self.assertIn('Missing action status field', errors)
        self.assertIn('Missing mail server or port', errors)
        self.assertIn('Missing mail from address', errors)
        self.assertIn('Missing mail subject', errors)
        self.assertIn('Missing message table headers', errors)

    def test_config_cc(self):
        uut = Configuration()

        # start with a baseline config
        config = SafeConfigParser()
        config.read('example/config.ini')

        config.remove_option(CSECT_EMAIL, 'cc')
        uut.update_config(config)
        self.assertEqual(None, uut.mail_cc)

        config.set(CSECT_EMAIL, 'cc', '')
        uut.update_config(config)
        self.assertEqual(None, uut.mail_cc)

        config.set(CSECT_EMAIL, 'cc', ',   ,    ,')
        uut.update_config(config)
        self.assertEqual(None, uut.mail_cc)

        mail1 = "fred@flintstone.com"
        mail2 = "barney@rubble.com"
        config.set(CSECT_EMAIL, 'cc', f"{mail1},   ,  {mail2}  ,")
        uut.update_config(config)
        self.assertEqual([mail1, mail2], uut.mail_cc)

    def test_config_table_headers(self):
        uut = Configuration()

        # start with a baseline config
        config = SafeConfigParser()
        config.read('example/config.ini')

        config.remove_option(CSECT_MSG, 'columns')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_headers)

        config.set(CSECT_MSG, 'columns', '')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_headers)

        config.set(CSECT_MSG, 'columns', ',   ,    ,')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_headers)

        hdr1 = "User"
        hdr2 = "Due Date"
        config.set(CSECT_MSG, 'columns', f"  , {hdr1},   ,  {hdr2}  ,")
        uut.update_config(config)
        self.assertEqual([hdr1, hdr2], uut.msg_table_headers)

    def test_config_table_align(self):
        uut = Configuration()

        # start with a baseline config
        config = SafeConfigParser()
        config.read('example/config.ini')

        config.remove_option(CSECT_MSG, 'align')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_align)

        config.set(CSECT_MSG, 'align', '')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_align)

        config.set(CSECT_MSG, 'align', ',   ,    ,')
        uut.update_config(config)
        self.assertEqual(None, uut.msg_table_align)

        hdr1 = "User"
        hdr2 = "Due Date"
        config.set(CSECT_MSG, 'align', f"  , {hdr1},   ,  {hdr2}: l  ,")
        self.assertRaises(ValueError, lambda: uut.update_config(config))

        config.set(CSECT_MSG, 'align', f"  , {hdr1}: r,   ,  {hdr2}: l  ,")
        uut.update_config(config)
        self.assertEqual({hdr1: 'r', hdr2: 'l'}, uut.msg_table_align)
