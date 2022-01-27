import datetime
import math
import time

import requests_mock

from exchangelib.configuration import Configuration
from exchangelib.credentials import Credentials
from exchangelib.protocol import FailFast, FaultTolerance
from exchangelib.transport import AUTH_TYPE_MAP, NTLM
from exchangelib.version import Build, Version

from .common import TimedTestCase, get_random_string


class ConfigurationTest(TimedTestCase):
    def test_init(self):
        with self.assertRaises(TypeError) as e:
            Configuration(credentials="foo")
        self.assertEqual(
            e.exception.args[0], "'credentials' 'foo' must be of type <class 'exchangelib.credentials.BaseCredentials'>"
        )
        with self.assertRaises(ValueError) as e:
            Configuration(credentials=None, auth_type=NTLM)
        self.assertEqual(e.exception.args[0], "Auth type 'NTLM' was detected but no credentials were provided")
        with self.assertRaises(AttributeError) as e:
            Configuration(server="foo", service_endpoint="bar")
        self.assertEqual(e.exception.args[0], "Only one of 'server' or 'service_endpoint' must be provided")
        with self.assertRaises(ValueError) as e:
            Configuration(auth_type="foo")
        self.assertEqual(e.exception.args[0], f"'auth_type' 'foo' must be one of {sorted(AUTH_TYPE_MAP)}")
        with self.assertRaises(TypeError) as e:
            Configuration(version="foo")
        self.assertEqual(e.exception.args[0], "'version' 'foo' must be of type <class 'exchangelib.version.Version'>")
        with self.assertRaises(TypeError) as e:
            Configuration(retry_policy="foo")
        self.assertEqual(
            e.exception.args[0], "'retry_policy' 'foo' must be of type <class 'exchangelib.protocol.RetryPolicy'>"
        )
        with self.assertRaises(TypeError) as e:
            Configuration(max_connections="foo")
        self.assertEqual(e.exception.args[0], "'max_connections' 'foo' must be of type <class 'int'>")
        self.assertEqual(Configuration().server, None)  # Test that property works when service_endpoint is None

    def test_magic(self):
        config = Configuration(
            server="example.com",
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NTLM,
            version=Version(build=Build(15, 1, 2, 3), api_version="foo"),
        )
        # Just test that these work
        str(config)
        repr(config)

    @requests_mock.mock()  # Just to make sure we don't make any requests
    def test_hardcode_all(self, m):
        # Test that we can hardcode everything without having a working server. This is useful if neither tasting or
        # guessing missing values works.
        Configuration(
            server="example.com",
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NTLM,
            version=Version(build=Build(15, 1, 2, 3), api_version="foo"),
        )

    def test_fail_fast_back_off(self):
        # Test that FailFast does not support back-off logic
        c = FailFast()
        self.assertIsNone(c.back_off_until)
        with self.assertRaises(AttributeError):
            c.back_off_until = 1

    def test_service_account_back_off(self):
        # Test back-off logic in FaultTolerance
        sa = FaultTolerance()

        # Initially, the value is None
        self.assertIsNone(sa.back_off_until)

        # Test a non-expired back off value
        in_a_while = datetime.datetime.now() + datetime.timedelta(seconds=10)
        sa.back_off_until = in_a_while
        self.assertEqual(sa.back_off_until, in_a_while)

        # Test an expired back off value
        sa.back_off_until = datetime.datetime.now()
        time.sleep(0.001)
        self.assertIsNone(sa.back_off_until)

        # Test the back_off() helper
        sa.back_off(10)
        # This is not a precise test. Assuming fast computers, there should be less than 1 second between the two lines.
        self.assertEqual(int(math.ceil((sa.back_off_until - datetime.datetime.now()).total_seconds())), 10)

        # Test expiry
        sa.back_off(0)
        time.sleep(0.001)
        self.assertIsNone(sa.back_off_until)

        # Test default value
        sa.back_off(None)
        self.assertEqual(int(math.ceil((sa.back_off_until - datetime.datetime.now()).total_seconds())), 60)
