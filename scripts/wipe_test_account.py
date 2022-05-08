from tests.common import EWSTest

from exchangelib.protocol import BaseProtocol
BaseProtocol.TIMEOUT = 300  # Seconds
t = EWSTest()
t.setUpClass()
t.wipe_test_account()
