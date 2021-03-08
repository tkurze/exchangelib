from exchangelib.errors import ErrorInvalidSubscription
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

        # Test that we see a delete event
        i1_id = i1.id
        i1.delete()
        changes = list(test_folder.sync_items())
        self.assertEqual(len(changes), 1)
        change_type, i = changes[0]
        self.assertEqual(change_type, 'delete')
        self.assertEqual(i.id, i1_id)

    @staticmethod
    def _filter_events(events, event_cls, item_id):
        return [e for e in events if isinstance(e, event_cls) and e.event_type == CreatedEvent.ITEM
                and e.item_id.id == item_id]

    def test_pull_notifications(self):
        # Test that we can create a pull subscription, make changes and see the events by calling .get_events()
        test_folder = self.account.drafts
        subscription_id, watermark = test_folder.subscribe_to_pull()
        notifications = list(test_folder.get_events(subscription_id, watermark))
        self.assertEqual(len(notifications), 1)
        notification = notifications[0]
        self.assertEqual(len(notification.events), 1)
        status_event = notification.events[0]
        self.assertIsInstance(status_event, StatusEvent)
        # Set the new watermark
        watermark = status_event.watermark

        # Test that we see a create event
        i1 = self.get_test_item(folder=test_folder).save()
        import time
        time.sleep(5)  # TODO: implement affinity
        notifications = list(test_folder.get_events(subscription_id, watermark))
        self.assertEqual(len(notifications), 1)
        notification = notifications[0]
        created_events = self._filter_events(notification.events, CreatedEvent, i1.id)
        self.assertEqual(len(created_events), 1)
        created_event = created_events[0]
        self.assertIsInstance(created_event, CreatedEvent)
        self.assertEqual(created_event.item_id.id, i1.id)
        # Set the new watermark
        watermark = notification.events[-1].watermark

        # Test that we see an update event
        i1.subject = get_random_string(8)
        i1.save(update_fields=['subject'])
        import time
        time.sleep(5)  # TODO: implement affinity
        notifications = list(test_folder.get_events(subscription_id, watermark))
        self.assertEqual(len(notifications), 1)
        notification = notifications[0]
        modified_events = self._filter_events(notification.events, ModifiedEvent, i1.id)
        self.assertEqual(len(modified_events), 1)
        modified_event = modified_events[0]
        self.assertIsInstance(modified_event, ModifiedEvent)
        self.assertEqual(modified_event.item_id.id, i1.id)
        # Set the new watermark
        watermark = notification.events[-1].watermark

        # Test that we see a delete event
        i1_id = i1.id
        i1.delete()
        import time
        time.sleep(5)  # TODO: implement affinity
        notifications = list(test_folder.get_events(subscription_id, watermark))
        self.assertEqual(len(notifications), 1)
        notification = notifications[0]
        deleted_events = self._filter_events(notification.events, DeletedEvent, i1_id)
        self.assertEqual(len(deleted_events), 1)
        deleted_event = deleted_events[0]
        self.assertIsInstance(deleted_event, DeletedEvent)
        self.assertEqual(deleted_event.item_id.id, i1_id)
        # Set the new watermark
        watermark = notification.events[-1].watermark

        test_folder.unsubscribe(subscription_id)
