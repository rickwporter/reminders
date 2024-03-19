import unittest

from reminders import SafeConfigParser


class TestConfig(unittest.TestCase):
    def test_config_empty(self):
        uut = SafeConfigParser()
        uut.read_string("")
        self.assertEqual('', uut.get("foo", "bar"))

    def test_config_basic(self):
        uut = SafeConfigParser()
        uut.read_string("[sna]\nfoo = bar")
        self.assertEqual('bar', uut.get('sna', 'foo'))
        self.assertEqual('', uut.get('sna', 'bar'))  # section exists
        self.assertEqual('', uut.get('blah', 'bar'))  # section does NOT exists

    def test_config_hash(self):
        uut = SafeConfigParser()
        uut.read_string("[sna]\nfoo = #bar")

        # we get back the hash value
        self.assertEqual('#bar', uut.get('sna', 'foo'))
