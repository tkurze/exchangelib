import pickle

from exchangelib.account import Identity
from exchangelib.credentials import Credentials, OAuth2Credentials, OAuth2AuthorizationCodeCredentials

from .common import TimedTestCase


class CredentialsTest(TimedTestCase):
    def test_hash(self):
        # Test that we can use credentials as a dict key
        self.assertEqual(hash(Credentials('a', 'b')), hash(Credentials('a', 'b')))
        self.assertNotEqual(hash(Credentials('a', 'b')), hash(Credentials('a', 'a')))
        self.assertNotEqual(hash(Credentials('a', 'b')), hash(Credentials('b', 'b')))

    def test_equality(self):
        self.assertEqual(Credentials('a', 'b'), Credentials('a', 'b'))
        self.assertNotEqual(Credentials('a', 'b'), Credentials('a', 'a'))
        self.assertNotEqual(Credentials('a', 'b'), Credentials('b', 'b'))

    def test_type(self):
        self.assertEqual(Credentials('a', 'b').type, Credentials.UPN)
        self.assertEqual(Credentials('a@example.com', 'b').type, Credentials.EMAIL)
        self.assertEqual(Credentials('a\\n', 'b').type, Credentials.DOMAIN)

    def test_pickle(self):
        # Test that we can pickle, hash, repr, str and compare various credentials types
        for o in (
            Identity('XXX', 'YYY', 'ZZZ', 'WWW'),
            Credentials('XXX', 'YYY'),
            OAuth2Credentials('XXX', 'YYY', 'ZZZZ'),
            OAuth2Credentials('XXX', 'YYY', 'ZZZZ', identity=Identity('AAA')),
            OAuth2AuthorizationCodeCredentials(client_id='WWW', client_secret='XXX'),
            OAuth2AuthorizationCodeCredentials(
                client_id='WWW', client_secret='XXX', authorization_code='YYY', access_token={'access_token': 'ZZZ'},
                tenant_id='ZZZ', identity=Identity('AAA')
            ),
        ):
            with self.subTest(o=o):
                pickled_o = pickle.dumps(o)
                unpickled_o = pickle.loads(pickled_o)
                self.assertIsInstance(unpickled_o, type(o))
                self.assertEqual(o, unpickled_o)
                self.assertEqual(hash(o), hash(unpickled_o))
                self.assertEqual(repr(o), repr(unpickled_o))
                self.assertEqual(str(o), str(unpickled_o))
