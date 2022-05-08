from exchangelib.protocol import BaseProtocol
from tests.common import EWSTest

BaseProtocol.TIMEOUT = 300  # Seconds
t = EWSTest()
t.setUpClass()
t.wipe_test_account()
