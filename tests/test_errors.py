from exchangelib.errors import EWSError

from .common import TimedTestCase


class ErrorTest(TimedTestCase):
    def test_hash(self):
        e = EWSError("foo")
        self.assertEqual(hash(e), hash("foo"))
