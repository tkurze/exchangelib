from .common import EWSTest
from exchangelib.errors import ErrorInvalidSubscription


class SyncTest(EWSTest):
    def test_pull_subscribe(self):
        subscription_id, watermark = self.account.inbox.subscribe_to_pull()
        self.assertIsNotNone(subscription_id)
        self.assertIsNotNone(watermark)
        self.account.inbox.unsubscribe(subscription_id)

    def test_push_subscribe(self):
        subscription_id, watermark = self.account.inbox.subscribe_to_push(callback_url='https://example.com/foo')
        self.assertIsNotNone(subscription_id)
        self.assertIsNotNone(watermark)
        with self.assertRaises(ErrorInvalidSubscription):
            self.account.inbox.unsubscribe(subscription_id)

    def test_streaming_subscribe(self):
        subscription_id = self.account.inbox.subscribe_to_streaming()
        self.assertIsNotNone(subscription_id)
        self.account.inbox.unsubscribe(subscription_id)
