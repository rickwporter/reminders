import unittest

from typing import Dict
from typing import List
from typing import Tuple
from unittest.mock import call
from unittest.mock import patch

from datetime import datetime
from reminders import AmbiguousUser
from reminders import Format
from reminders import MissingUser
from reminders import Reminders
from reminders import ROW_HEADER


class TestReminders(unittest.TestCase):
    def test_reminders_sheet_to_dict_errors(self):
        uut = Reminders()
        self.assertRaises(FileNotFoundError, lambda: uut.sheet_to_dict('foo', 'bar'))
        self.assertRaises(ValueError, lambda: uut.sheet_to_dict('example/bedrock.xlsx', 'bar'))

    def test_reminders_sheet_to_dict_example_action(self):
        uut = Reminders()
        items = uut.sheet_to_dict('example/bedrock.xlsx', 'Actions')
        self.assertEqual(4, len(items))
        first = items[0]
        self.assertEqual('SG1', first.get('ID'))
        self.assertEqual('Fred', first.get('User'))
        self.assertEqual('Break big rocks into small ones', first.get('Action'))
        self.assertEqual('Open', first.get('Status'))
        self.assertEqual('Far in the future', first.get('Notes'))
        self.assertEqual('This column is NOT included in the example', first.get('Extra'))
        self.assertEqual('Actions:1', first.get('_row'))
        due = first.get('Due Date')
        self.assertTrue(isinstance(due, datetime))
        self.assertEqual(datetime(year=2030, month=3, day=24), due)

    def test_reminders_sheet_to_dict_example_user(self):
        uut = Reminders()
        items = uut.sheet_to_dict('example/bedrock.xlsx', 'Users')
        self.assertEqual(4, len(items))
        keys = set(['User', 'First', 'Email', 'Aliases', ROW_HEADER])
        first = items[0]
        self.assertEqual(keys, set(first.keys()))
        self.assertEqual('Fred Flintstone', first.get('User'))
        self.assertEqual('Fred', first.get('First'))
        self.assertEqual('fred@slaterockandgravel.com', first.get('Email'))
        self.assertEqual('FF', first.get('Aliases'))
        self.assertEqual('Users:1', first.get('_row'))

    def test_reminders_find_user(self):
        uut = Reminders()
        users = uut.sheet_to_dict('example/bedrock.xlsx', 'Users')
        self.assertEqual(4, len(users))
        fred = users[0]  # Fred is the first row
        wilma = users[1]
        barney = users[2]
        betty = users[3]

        # find Fred
        self.assertEqual(fred, uut.find_user(users, 'Fred'))
        self.assertEqual(fred, uut.find_user(users, 'FF'))  # aliases
        self.assertEqual(fred, uut.find_user(users, 'ff'))  # aliases, lowercase
        self.assertEqual(fred, uut.find_user(users, 'FreD@slate'))  # email case insensitive
        self.assertEqual(fred, uut.find_user(users, 'Fred flint'))  # user case insensitive

        # find Wilma
        self.assertEqual(wilma, uut.find_user(users, 'Wilma F'))
        self.assertEqual(wilma, uut.find_user(users, 'wf'))  # aliases
        self.assertEqual(wilma, uut.find_user(users, 'wilma.Fl'))  # email case insensitive

        # find Barney
        self.assertEqual(barney, uut.find_user(users, 'Barney R'))
        self.assertEqual(barney, uut.find_user(users, 'BR'))  # aliases
        self.assertEqual(barney, uut.find_user(users, 'barney@'))  # email case insensitive

        # find Betty
        self.assertEqual(betty, uut.find_user(users, 'Betty R'))
        self.assertEqual(betty, uut.find_user(users, 'betty.rubble@'))  # email case insensitive

        # No such finding is NOT an error
        self.assertIsNone(uut.find_user(users, 'Bambam'))

        self.assertRaises(AmbiguousUser, lambda: uut.find_user(users, 'slate'))
        self.assertRaises(AmbiguousUser, lambda: uut.find_user(users, 'flintstone'))

    def test_reminders_config_example(self):
        uut = Reminders()
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

    def test_reminders_config_empty(self):
        uut = Reminders()
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

    def findActions(self, correlated: List[Tuple], firstname: str) -> List[Dict]:
        for (user, userActions) in correlated:
            if user.get('First') == firstname:
                return userActions
        return []

    def test_reminders_correlate_example(self):
        uut = Reminders()
        uut.parse_config('example/config.ini')  # correlation needs the fields initialized
        users = uut.sheet_to_dict('example/bedrock.xlsx', 'Users')
        actions = uut.sheet_to_dict('example/bedrock.xlsx', 'Actions')
        correlated = uut.correlate(users, actions)
        self.assertEqual(len(users), len(correlated))

        items = self.findActions(correlated, 'Fred')
        self.assertEqual(
            set(['Actions:1', 'Actions:2', 'Actions:3']),
            set([_.get(ROW_HEADER) for _ in items])
        )

        items = self.findActions(correlated, 'Barney')
        self.assertEqual(
            set(['Actions:2', 'Actions:4']),
            set([_.get(ROW_HEADER) for _ in items])
        )

        items = self.findActions(correlated, 'Betty')
        self.assertEqual(
            set(['Actions:4']),
            set([_.get(ROW_HEADER) for _ in items])
        )

        items = self.findActions(correlated, 'Wilma')
        self.assertEqual(
            set(['Actions:3']),
            set([_.get(ROW_HEADER) for _ in items])
        )

        # modify a user action to get a MissingUser exception
        actions[2].update({'User': 'Bam Bam'})
        self.assertRaises(MissingUser, lambda: uut.correlate(users, actions))

    def read_text(self, filename: str) -> str:
        with open(filename) as fp:
            return fp.read()

    def test_reminders_table(self):
        uut = Reminders()
        uut.parse_config('example/config.ini')  # table creation needs the fields initialized
        actions = uut.sheet_to_dict('example/bedrock.xlsx', 'Actions')

        expected = self.read_text('resources/all_actions.html')
        self.assertEqual(expected, uut._create_table(actions, Format.HTML))
        expected = self.read_text('resources/all_actions.txt')
        self.assertEqual(expected, uut._create_table(actions, Format.TEXT))
        expected = self.read_text('resources/all_actions.json')
        self.assertEqual(expected, uut._create_table(actions, Format.JSON))
        expected = self.read_text('resources/all_actions.csv').replace('\n', '\r\n')
        self.assertEqual(expected, uut._create_table(actions, Format.CSV))

    def test_reminders_validate_users(self):
        uut = Reminders()
        uut.parse_config('example/config.ini')  # correlation needs the fields initialized
        users = uut.sheet_to_dict('example/bedrock.xlsx', 'Users')

        self.assertEqual([], uut.validate_users(users))

        users[0].update({uut.hdr_email: ''})
        users[3].pop(uut.hdr_email, None)
        errors = uut.validate_users(users)
        self.assertEqual(2, len(errors))
        self.assertIn('Fred Flintstone (Users:1) error(s): missing email', errors)
        self.assertIn('Betty Rubble (Users:4) error(s): missing email', errors)

        # change template to catch everyone
        uut.msg_preamble = '{Foo}'
        uut.msg_close = '{Bar}'

        errors = uut.validate_users(users)
        self.assertEqual(4, len(errors))
        self.assertIn('Fred Flintstone (Users:1) error(s): missing email, missing email fields Bar/Foo', errors)
        self.assertIn('Wilma Flintstone (Users:2) error(s): missing email fields Bar/Foo', errors)
        self.assertIn('Barney Rubble (Users:3) error(s): missing email fields Bar/Foo', errors)
        self.assertIn('Betty Rubble (Users:4) error(s): missing email, missing email fields Bar/Foo', errors)

    def test_reminders_validate_actions(self):
        uut = Reminders()
        uut.parse_config('example/config.ini')  # correlation needs the fields initialized
        actions = uut.sheet_to_dict('example/bedrock.xlsx', 'Actions')

        self.assertEqual([], uut.validate_actions(actions))

        actions[0].update({uut.hdr_user: ''})
        actions[3].pop(uut.hdr_user, None)
        actions[3].update({uut.hdr_due: None})
        errors = uut.validate_actions(actions)
        self.assertEqual(2, len(errors))
        self.assertIn('SG1 (Actions:1) error(s): missing assignment', errors)
        self.assertIn('Rubble1 (Actions:4) error(s): missing assignment, missing due date', errors)

        # change template to catch all actions
        uut.msg_table_headers = ['Sna', 'Foo']
        errors = uut.validate_actions(actions)
        self.assertEqual(4, len(errors))
        self.assertIn('SG1 (Actions:1) error(s): missing assignment, missing table fields Foo/Sna', errors)
        self.assertIn('SG2 (Actions:2) error(s): missing table fields Foo/Sna', errors)
        self.assertIn('FS1 (Actions:3) error(s): missing table fields Foo/Sna', errors)
        self.assertIn('Rubble1 (Actions:4) error(s): missing assignment, missing due date, missing table fields Foo/Sna', errors)  # noqa: E501

    @patch('reminders.Reminders.print')
    @patch('smtplib.SMTP.sendmail')
    @patch('smtplib.SMTP.login')
    @patch('smtplib.SMTP.starttls')
    def test_reminders_run_example(self, mock_starttls, mock_login, mock_send, mock_print):
        uut = Reminders()
        uut.mail_password = 'abc123'  # avoid prompting

        args = [
            '-c',
            'example/config.ini',
            # NOTE: must specify filename, since config.ini assumes it is in same directory as config.ini
            '-s',
            'example/bedrock.xlsx',
        ]
        result = uut.run(args)
        self.assertEqual(0, result)
        self.assertEqual(1, mock_starttls.call_count)
        self.assertEqual(1, mock_login.call_count)
        mock_login.assert_called_once_with(uut.mail_from, uut.mail_password)
        self.assertEqual(3, mock_send.call_count)
        self.assertEqual(4, mock_print.call_count)
        print_calls = [
            call('Sending emails about items due in the next 14 days:'),
            call('    Barney Rubble: 2'),
            call('    Betty Rubble: 1'),
            call('    Fred Flintstone: 1'),
        ]
        self.assertEqual(print_calls, mock_print.call_args_list)

    @patch('reminders.Reminders.print')
    def test_reminders_run_no_config(self, mock_print):
        uut = Reminders()
        uut.mail_password = 'abc123'  # avoid prompting

        args = [
            '-s',
            'example/bedrock.xlsx',
        ]
        result = uut.run(args)
        self.assertEqual(5, result)
        mock_print.assert_called_once_with('Configuration errors:\n\tMissing spreadsheet tab name for actions\n\tMissing spreadsheet tab name for users\n\tMissing user/action user-id field\n\tMissing user email field\n\tMissing action identifier field\n\tMissing action due date field\n\tMissing action status field\n\tMissing mail server or port\n\tMissing mail from address\n\tMissing mail subject\n\tMissing message table headers')  # noqa: E501

    @patch('reminders.Reminders.print')
    def test_reminders_run_no_spreadsheet(self, mock_print):
        uut = Reminders()
        uut.mail_password = 'abc123'  # avoid prompting

        args = [
            '-c',
            'example/config.ini',
        ]
        result = uut.run(args)
        self.assertEqual(1, result)
        mock_print.assert_called_once_with('bedrock.xlsx is not a file!')

    @patch('reminders.Reminders.print')
    @patch('smtplib.SMTP.sendmail')
    @patch('smtplib.SMTP.login')
    @patch('smtplib.SMTP.starttls')
    def test_reminders_run_person_filter(self, mock_starttls, mock_login, mock_send, mock_print):
        uut = Reminders()
        uut.mail_password = 'abc123'  # avoid prompting

        args = [
            '-c',
            'example/config.ini',
            # NOTE: must specify filename, since config.ini assumes it is in same directory as config.ini
            '-s',
            'example/bedrock.xlsx',
            '-p',
            'fred',
            '-d',
            '30',
        ]
        result = uut.run(args)
        self.assertEqual(0, result)
        self.assertEqual(1, mock_starttls.call_count)
        self.assertEqual(1, mock_login.call_count)
        mock_login.assert_called_once_with(uut.mail_from, uut.mail_password)
        self.assertEqual(1, mock_send.call_count)
        self.assertEqual(2, mock_print.call_count)
        print_calls = [
            call('Sending emails about items due in the next 30 days:'),
            call('    Fred Flintstone: 1'),
        ]
        self.assertEqual(print_calls, mock_print.call_args_list)

    @patch('reminders.Reminders.print')
    def test_reminders_run_missing_person_filter(self, mock_print):
        uut = Reminders()
        uut.mail_password = 'abc123'  # avoid prompting

        args = [
            '-c',
            'example/config.ini',
            # NOTE: must specify filename, since config.ini assumes it is in same directory as config.ini
            '-s',
            'example/bedrock.xlsx',
            '-p',
            'Pebbles',
        ]
        result = uut.run(args)
        self.assertEqual(0, result)
        mock_print.assert_called_once_with('No open user actions found in example/bedrock.xlsx for Pebbles in the next 14 days')  # noqa: E501
