import time

from exchangelib.errors import ErrorInvalidSubscription, ErrorSubscriptionNotFound
from exchangelib.folders import Inbox
from exchangelib.items import Message
from exchangelib.properties import StatusEvent, CreatedEvent, ModifiedEvent, DeletedEvent

from .test_basics import BaseItemTest
from ..common import get_random_string


class SyncTest(BaseItemTest):
    TEST_FOLDER = 'inbox'
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_pull_subscribe(self):
        with self.account.inbox.pull_subscription() as (subscription_id, watermark):
            self.assertIsNotNone(subscription_id)
            self.assertIsNotNone(watermark)
        # Context manager already unsubscribed us
        with self.assertRaises(ErrorSubscriptionNotFound):
            self.account.inbox.unsubscribe(subscription_id)

    def test_push_subscribe(self):
        with self.account.inbox.push_subscription(
                callback_url='https://example.com/foo'
        ) as (subscription_id, watermark):
            self.assertIsNotNone(subscription_id)
            self.assertIsNotNone(watermark)
        with self.assertRaises(ErrorInvalidSubscription):
            self.account.inbox.unsubscribe(subscription_id)

    def test_streaming_subscribe(self):
        with self.account.inbox.streaming_subscription() as subscription_id:
            self.assertIsNotNone(subscription_id)
        # Context manager already unsubscribed us
        with self.assertRaises(ErrorSubscriptionNotFound):
            self.account.inbox.unsubscribe(subscription_id)

    def test_sync_folder_hierarchy(self):
        test_folder = self.get_test_folder().save()

        # Test that folder_sync_state is set after calling sync_hierarchy
        self.assertIsNone(test_folder.folder_sync_state)
        list(test_folder.sync_hierarchy())
        self.assertIsNotNone(test_folder.folder_sync_state)

        # Test that we see a create event
        f1 = self.FOLDER_CLASS(parent=test_folder, name=get_random_string(8)).save()
        changes = list(test_folder.sync_hierarchy())
        self.assertEqual(len(changes), 1)
        change_type, f = changes[0]
        self.assertEqual(change_type, 'create')
        self.assertEqual(f.id, f1.id)

        # Test that we see an update event
        f1.name = get_random_string(8)
        f1.save(update_fields=['name'])
        changes = list(test_folder.sync_hierarchy())
        self.assertEqual(len(changes), 1)
        change_type, f = changes[0]
        self.assertEqual(change_type, 'update')
        self.assertEqual(f.id, f1.id)

        # Test that we see a delete event
        f1_id = f1.id
        f1.delete()
        changes = list(test_folder.sync_hierarchy())
        self.assertEqual(len(changes), 1)
        change_type, f = changes[0]
        self.assertEqual(change_type, 'delete')
        self.assertEqual(f.id, f1_id)

    def test_sync_folder_items(self):
        test_folder = self.get_test_folder().save()

        # Test that item_sync_state is set after calling sync_hierarchy
        self.assertIsNone(test_folder.item_sync_state)
        list(test_folder.sync_items())
        self.assertIsNotNone(test_folder.item_sync_state)

        # Test that we see a create event
        i1 = self.get_test_item(folder=test_folder).save()
        changes = list(test_folder.sync_items())
        self.assertEqual(len(changes), 1)
        change_type, i = changes[0]
        self.assertEqual(change_type, 'create')
        self.assertEqual(i.id, i1.id)

        # Test that we see an update event
        i1.subject = get_random_string(8)
        i1.save(update_fields=['subject'])
        changes = list(test_folder.sync_items())
        self.assertEqual(len(changes), 1)
        change_type, i = changes[0]
        self.assertEqual(change_type, 'update')
        self.assertEqual(i.id, i1.id)

        # Test that we see a read_flag_change event
        i1.is_read = not i1.is_read
        i1.save(update_fields=['is_read'])
        changes = list(test_folder.sync_items())
        self.assertEqual(len(changes), 1)
        change_type, (i, read_state) = changes[0]
        self.assertEqual(change_type, 'read_flag_change')
        self.assertEqual(i.id, i1.id)
        self.assertEqual(read_state, i1.is_read)

        # Test that we see a delete event
        i1_id = i1.id
        i1.delete()
        changes = list(test_folder.sync_items())
        self.assertEqual(len(changes), 1)
        change_type, i = changes[0]
        self.assertEqual(change_type, 'delete')
        self.assertEqual(i.id, i1_id)

    def _filter_events(self, notifications, event_cls, item_id):
        events = []
        watermark = None
        for notification in notifications:
            for e in notification.events:
                watermark = e.watermark
                if not isinstance(e, event_cls):
                    continue
                if item_id is None:
                    events.append(e)
                    continue
                if e.event_type == event_cls.ITEM and e.item_id.id == item_id:
                    events.append(e)
        self.assertEqual(len(events), 1)
        event = events[0]
        self.assertIsInstance(event, event_cls)
        return event, watermark

    def test_pull_notifications(self):
        # Test that we can create a pull subscription, make changes and see the events by calling .get_events()
        test_folder = self.account.drafts
        with test_folder.pull_subscription() as (subscription_id, watermark):
            notifications = list(test_folder.get_events(subscription_id, watermark))
            _, watermark = self._filter_events(notifications, StatusEvent, None)

            # Test that we see a create event
            i1 = self.get_test_item(folder=test_folder).save()
            time.sleep(5)  # TODO: For some reason, events do not trigger instantly
            notifications = list(test_folder.get_events(subscription_id, watermark))
            created_event, watermark = self._filter_events(notifications, CreatedEvent, i1.id)
            self.assertEqual(created_event.item_id.id, i1.id)

            # Test that we see an update event
            i1.subject = get_random_string(8)
            i1.save(update_fields=['subject'])
            time.sleep(5)  # TODO: For some reason, events do not trigger instantly
            notifications = list(test_folder.get_events(subscription_id, watermark))
            modified_event, watermark = self._filter_events(notifications, ModifiedEvent, i1.id)
            self.assertEqual(modified_event.item_id.id, i1.id)

            # Test that we see a delete event
            i1_id = i1.id
            i1.delete()
            time.sleep(5)  # TODO: For some reason, events do not trigger instantly
            notifications = list(test_folder.get_events(subscription_id, watermark))
            deleted_event, watermark = self._filter_events(notifications, DeletedEvent, i1_id)
            self.assertEqual(deleted_event.item_id.id, i1_id)

    def test_streaming_notifications(self):
        # Test that we can create a streaming subscription, make changes and see the events by calling
        # .get_streaming_events()
        test_folder = self.account.drafts
        with test_folder.streaming_subscription() as subscription_id:
            # Test that we see a create event
            i1 = self.get_test_item(folder=test_folder).save()
            # 1 minute connection timeout
            notifications = list(test_folder.get_streaming_events(
                subscription_id, connection_timeout=1, max_notifications_returned=1
            ))
            created_event, _ = self._filter_events(notifications, CreatedEvent, i1.id)
            self.assertEqual(created_event.item_id.id, i1.id)

            # Test that we see an update event
            i1.subject = get_random_string(8)
            i1.save(update_fields=['subject'])
            # 1 minute connection timeout
            notifications = list(test_folder.get_streaming_events(
                subscription_id, connection_timeout=1, max_notifications_returned=1
            ))
            modified_event, _ = self._filter_events(notifications, ModifiedEvent, i1.id)
            self.assertEqual(modified_event.item_id.id, i1.id)

            # Test that we see a delete event
            i1_id = i1.id
            i1.delete()
            # 1 minute connection timeout
            notifications = list(test_folder.get_streaming_events(
                subscription_id, connection_timeout=1, max_notifications_returned=1
            ))
            deleted_event, _ = self._filter_events(notifications, DeletedEvent, i1_id)
            self.assertEqual(deleted_event.item_id.id, i1_id)

    def test_streaming_with_other_calls(self):
        # Test that we can call other EWS operations while we have a streaming subscription open
        test_folder = self.account.drafts

        # Test calling GetItem while the streaming connection is still open. We need to bump the
        # connection count because the default count is 1 but we need 2 connections.
        self.account.protocol._session_pool_maxsize += 1
        self.account.protocol.increase_poolsize()
        try:
            with test_folder.streaming_subscription() as subscription_id:
                i1 = self.get_test_item(folder=test_folder).save()
                for notification in test_folder.get_streaming_events(
                    subscription_id, connection_timeout=1, max_notifications_returned=1
                ):
                    for e in notification.events:
                        if isinstance(e, CreatedEvent) and e.event_type == CreatedEvent.ITEM \
                                and e.item_id.id == i1.id:
                            test_folder.all().only('id').get(id=e.item_id.id)
        finally:
            self.account.protocol.decrease_poolsize()
            self.account.protocol._session_pool_maxsize -= 1
