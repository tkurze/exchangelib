import logging
import random
import sys
from unittest import TestLoader, TestSuite
import unittest.util

from exchangelib.util import PrettyXmlHandler


class RandomTestSuite(TestSuite):
    def __iter__(self):
        tests = list(super().__iter__())
        random.shuffle(tests)
        return iter(tests)


# Execute test classes in random order
TestLoader.suiteClass = RandomTestSuite
# Execute test methods in random order within each test class
TestLoader.sortTestMethodsUsing = lambda _, x, y: random.choice((1, -1))

# Always show full repr() output for object instances in unittest error messages
unittest.util._MAX_LENGTH = 2000

if '-v' in sys.argv:
    logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])
else:
    logging.basicConfig(level=logging.CRITICAL)
