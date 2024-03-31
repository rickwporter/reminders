import unittest

from datetime import datetime
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
        self.assertEqual(fred, uut.findUser(users, 'Fred'))
        self.assertEqual(fred, uut.findUser(users, 'FF'))  # aliases
        self.assertEqual(fred, uut.findUser(users, 'ff'))  # aliases, lowercase
        self.assertEqual(fred, uut.findUser(users, 'FreD@slate'))  # email case insensitive
        self.assertEqual(fred, uut.findUser(users, 'Fred flint'))  # user case insensitive

        # TODO: fix duplicate error -- should be an error
        self.assertEqual(fred, uut.findUser(users, 'slate'))
        self.assertEqual(fred, uut.findUser(users, 'flintstone'))

        # find Wilma
        self.assertEqual(wilma, uut.findUser(users, 'Wilma F'))
        self.assertEqual(wilma, uut.findUser(users, 'wf'))  # aliases
        self.assertEqual(wilma, uut.findUser(users, 'wilma.Fl'))  # email case insensitive

        # find Barney
        self.assertEqual(barney, uut.findUser(users, 'Barney R'))
        self.assertEqual(barney, uut.findUser(users, 'BR'))  # aliases
        self.assertEqual(barney, uut.findUser(users, 'barney@'))  # email case insensitive

        # find Betty
        self.assertEqual(betty, uut.findUser(users, 'Betty R'))
        self.assertEqual(betty, uut.findUser(users, 'betty.rubble@'))  # email case insensitive
