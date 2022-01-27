import logging
import os
import random
import unittest.util
from unittest import TestLoader, TestSuite

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
# Make sure we're also random in multiprocess test runners
random.seed()

# Always show full repr() output for object instances in unittest error messages
unittest.util._MAX_LENGTH = 2000

if os.environ.get("DEBUG", "").lower() in ("1", "yes", "true"):
    logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])
else:
    logging.basicConfig(level=logging.CRITICAL)
