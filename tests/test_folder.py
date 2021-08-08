from exchangelib.errors import ErrorDeleteDistinguishedFolder, ErrorObjectTypeChanged, DoesNotExist, \
    MultipleObjectsReturned, ErrorItemSave, ErrorItemNotFound
from exchangelib.extended_properties import ExtendedProperty
from exchangelib.items import Message
from exchangelib.folders import Calendar, DeletedItems, Drafts, Inbox, Outbox, SentItems, JunkEmail, Messages, Tasks, \
    Contacts, Folder, RecipientCache, GALContacts, System, AllContacts, MyContactsExtended, Reminders, Favorites, \
    AllItems, ConversationSettings, Friends, RSSFeeds, Sharing, IMContactList, QuickContacts, Journal, Notes, \
    SyncIssues, MyContacts, ToDoSearch, FolderCollection, DistinguishedFolderId, Files, \
    DefaultFoldersChangeHistory, PassThroughSearchResults, SmsAndChatsSync, GraphAnalytics, Signal, \
    PdpProfileV2Secured, VoiceMail, FolderQuerySet, SingleFolderQuerySet, SHALLOW, RootOfHierarchy, Companies, \
    OrganizationalContacts, PeopleCentricConversationBuddies, PublicFoldersRoot
from exchangelib.properties import Mailbox, InvalidField, EffectiveRights, PermissionSet, CalendarPermission, UserId
from exchangelib.queryset import Q
from exchangelib.services import GetFolder

from .common import EWSTest, get_random_string, get_random_int, get_random_bool, get_random_datetime, get_random_bytes,\
    get_random_byte


def get_random_str_tuple(tuple_length, str_length):
    return tuple(get_random_string(str_length, spaces=False) for _ in range(tuple_length))


class FolderTest(EWSTest):
    def test_folders(self):
        for f in self.account.root.walk():
            if isinstance(f, System):
                # No access to system folder, apparently
                continue
            f.test_access()
        # Test shortcuts
        for f, cls in (
                (self.account.trash, DeletedItems),
                (self.account.drafts, Drafts),
                (self.account.inbox, Inbox),
                (self.account.outbox, Outbox),
                (self.account.sent, SentItems),
                (self.account.junk, JunkEmail),
                (self.account.contacts, Contacts),
                (self.account.tasks, Tasks),
                (self.account.calendar, Calendar),
        ):
            with self.subTest(f=f, cls=cls):
                self.assertIsInstance(f, cls)
                f.test_access()
                # Test item field lookup
                self.assertEqual(f.get_item_field_by_fieldname('subject').name, 'subject')
                with self.assertRaises(ValueError):
                    f.get_item_field_by_fieldname('XXX')

    def test_folder_failure(self):
        # Folders must have an ID
        with self.assertRaises(ValueError):
            self.account.root.get_folder(Folder())
        with self.assertRaises(ValueError):
            self.account.root.add_folder(Folder())
        with self.assertRaises(ValueError):
            self.account.root.update_folder(Folder())
        with self.assertRaises(ValueError):
            self.account.root.remove_folder(Folder())
        # Removing a non-existent folder is allowed
        self.account.root.remove_folder(Folder(id='XXX'))
        # Must be called on a distinguished folder class
        with self.assertRaises(ValueError):
            RootOfHierarchy.get_distinguished(self.account)
        with self.assertRaises(ValueError):
            self.account.root.get_default_folder(Folder)

    def test_public_folders_root(self):
        # Test account does not have a public folders root. Make a dummy query just to hit .get_children()
        self.assertGreaterEqual(
            len(list(PublicFoldersRoot(account=self.account, is_distinguished=True).get_children(self.account.inbox))),
            0,
        )

    def test_find_folders(self):
        folders = list(FolderCollection(account=self.account, folders=[self.account.root]).find_folders())
        self.assertGreater(len(folders), 40, sorted(f.name for f in folders))

    def test_find_folders_with_restriction(self):
        # Exact match
        folders = list(FolderCollection(account=self.account, folders=[self.account.root])
                       .find_folders(q=Q(name='Top of Information Store')))
        self.assertEqual(len(folders), 1, sorted(f.name for f in folders))
        # Startswith
        folders = list(FolderCollection(account=self.account, folders=[self.account.root])
                       .find_folders(q=Q(name__startswith='Top of ')))
        self.assertEqual(len(folders), 1, sorted(f.name for f in folders))
        # Wrong case
        folders = list(FolderCollection(account=self.account, folders=[self.account.root])
                       .find_folders(q=Q(name__startswith='top of ')))
        self.assertEqual(len(folders), 0, sorted(f.name for f in folders))
        # Case insensitive
        folders = list(FolderCollection(account=self.account, folders=[self.account.root])
                       .find_folders(q=Q(name__istartswith='top of ')))
        self.assertEqual(len(folders), 1, sorted(f.name for f in folders))

    def test_get_folders(self):
        folders = list(FolderCollection(account=self.account, folders=[self.account.root]).get_folders())
        self.assertEqual(len(folders), 1, sorted(f.name for f in folders))

        # Test that GetFolder can handle FolderId instances
        folders = list(FolderCollection(account=self.account, folders=[DistinguishedFolderId(
            id=Inbox.DISTINGUISHED_FOLDER_ID,
            mailbox=Mailbox(email_address=self.account.primary_smtp_address)
        )]).get_folders())
        self.assertEqual(len(folders), 1, sorted(f.name for f in folders))

    def test_get_folders_with_distinguished_id(self):
        # Test that we return an Inbox instance and not a generic Messages or Folder instance when we call GetFolder
        # with a DistinguishedFolderId instance with an ID of Inbox.DISTINGUISHED_FOLDER_ID.
        inbox = list(GetFolder(account=self.account).call(
            folders=[DistinguishedFolderId(
                id=Inbox.DISTINGUISHED_FOLDER_ID,
                mailbox=Mailbox(email_address=self.account.primary_smtp_address))
            ],
            shape='IdOnly',
            additional_fields=[],
        ))[0]
        self.assertIsInstance(inbox, Inbox)

    def test_folder_grouping(self):
        # If you get errors here, you probably need to fill out [folder class].LOCALIZED_NAMES for your locale.
        for f in self.account.root.walk():
            with self.subTest(f=f):
                if isinstance(f, (
                        Messages, DeletedItems, AllContacts, MyContactsExtended, Sharing, Favorites, SyncIssues,
                        MyContacts
                )):
                    self.assertEqual(f.folder_class, 'IPF.Note')
                elif isinstance(f, Companies):
                    self.assertEqual(f.folder_class, 'IPF.Contact.Company')
                elif isinstance(f, OrganizationalContacts):
                    self.assertEqual(f.folder_class, 'IPF.Contact.OrganizationalContacts')
                elif isinstance(f, PeopleCentricConversationBuddies):
                    self.assertEqual(f.folder_class, 'IPF.Contact.PeopleCentricConversationBuddies')
                elif isinstance(f, GALContacts):
                    self.assertEqual(f.folder_class, 'IPF.Contact.GalContacts')
                elif isinstance(f, RecipientCache):
                    self.assertEqual(f.folder_class, 'IPF.Contact.RecipientCache')
                elif isinstance(f, Contacts):
                    self.assertEqual(f.folder_class, 'IPF.Contact')
                elif isinstance(f, Calendar):
                    self.assertEqual(f.folder_class, 'IPF.Appointment')
                elif isinstance(f, (Tasks, ToDoSearch)):
                    self.assertEqual(f.folder_class, 'IPF.Task')
                elif isinstance(f, Reminders):
                    self.assertEqual(f.folder_class, 'Outlook.Reminder')
                elif isinstance(f, AllItems):
                    self.assertEqual(f.folder_class, 'IPF')
                elif isinstance(f, ConversationSettings):
                    self.assertEqual(f.folder_class, 'IPF.Configuration')
                elif isinstance(f, Files):
                    self.assertEqual(f.folder_class, 'IPF.Files')
                elif isinstance(f, Friends):
                    self.assertEqual(f.folder_class, 'IPF.Note')
                elif isinstance(f, RSSFeeds):
                    self.assertEqual(f.folder_class, 'IPF.Note.OutlookHomepage')
                elif isinstance(f, IMContactList):
                    self.assertEqual(f.folder_class, 'IPF.Contact.MOC.ImContactList')
                elif isinstance(f, QuickContacts):
                    self.assertEqual(f.folder_class, 'IPF.Contact.MOC.QuickContacts')
                elif isinstance(f, Journal):
                    self.assertEqual(f.folder_class, 'IPF.Journal')
                elif isinstance(f, Notes):
                    self.assertEqual(f.folder_class, 'IPF.StickyNote')
                elif isinstance(f, DefaultFoldersChangeHistory):
                    self.assertEqual(f.folder_class, 'IPM.DefaultFolderHistoryItem')
                elif isinstance(f, PassThroughSearchResults):
                    self.assertEqual(f.folder_class, 'IPF.StoreItem.PassThroughSearchResults')
                elif isinstance(f, SmsAndChatsSync):
                    self.assertEqual(f.folder_class, 'IPF.SmsAndChatsSync')
                elif isinstance(f, GraphAnalytics):
                    self.assertEqual(f.folder_class, 'IPF.StoreItem.GraphAnalytics')
                elif isinstance(f, Signal):
                    self.assertEqual(f.folder_class, 'IPF.StoreItem.Signal')
                elif isinstance(f, PdpProfileV2Secured):
                    self.assertEqual(f.folder_class, 'IPF.StoreItem.PdpProfileSecured')
                elif isinstance(f, VoiceMail):
                    self.assertEqual(f.folder_class, 'IPF.Note.Microsoft.Voicemail')
                else:
                    self.assertIn(f.folder_class, (None, 'IPF'), (f.name, f.__class__.__name__, f.folder_class))
                    self.assertIsInstance(f, Folder)

    def test_counts(self):
        # Test count values on a folder
        f = Folder(parent=self.account.inbox, name=get_random_string(16)).save()
        f.refresh()

        self.assertEqual(f.total_count, 0)
        self.assertEqual(f.unread_count, 0)
        self.assertEqual(f.child_folder_count, 0)
        # Create some items
        items = []
        for i in range(3):
            subject = 'Test Subject %s' % i
            item = Message(account=self.account, folder=f, is_read=False, subject=subject, categories=self.categories)
            item.save()
            items.append(item)
        # Refresh values and see that total_count and unread_count changes
        f.refresh()
        self.assertEqual(f.total_count, 3)
        self.assertEqual(f.unread_count, 3)
        self.assertEqual(f.child_folder_count, 0)
        for i in items:
            i.is_read = True
            i.save()
        # Refresh values and see that unread_count changes
        f.refresh()
        self.assertEqual(f.total_count, 3)
        self.assertEqual(f.unread_count, 0)
        self.assertEqual(f.child_folder_count, 0)
        self.bulk_delete(items)
        # Refresh values and see that total_count changes
        f.refresh()
        self.assertEqual(f.total_count, 0)
        self.assertEqual(f.unread_count, 0)
        self.assertEqual(f.child_folder_count, 0)
        # Create some subfolders
        subfolders = []
        for i in range(3):
            subfolders.append(Folder(parent=f, name=get_random_string(16)).save())
        # Refresh values and see that child_folder_count changes
        f.refresh()
        self.assertEqual(f.total_count, 0)
        self.assertEqual(f.unread_count, 0)
        self.assertEqual(f.child_folder_count, 3)
        for sub_f in subfolders:
            sub_f.delete()
        # Refresh values and see that child_folder_count changes
        f.refresh()
        self.assertEqual(f.total_count, 0)
        self.assertEqual(f.unread_count, 0)
        self.assertEqual(f.child_folder_count, 0)
        f.delete()

    def test_refresh(self):
        # Test that we can refresh folders
        for f in self.account.root.walk():
            with self.subTest(f=f):
                if isinstance(f, System):
                    # Can't refresh the 'System' folder for some reason
                    continue
                old_values = {}
                for field in f.FIELDS:
                    old_values[field.name] = getattr(f, field.name)
                    if field.name in ('account', 'id', 'changekey', 'parent_folder_id'):
                        # These are needed for a successful refresh()
                        continue
                    if field.is_read_only:
                        continue
                    setattr(f, field.name, self.random_val(field))
                f.refresh()
                for field in f.FIELDS:
                    if field.name == 'changekey':
                        # folders may change while we're testing
                        continue
                    if field.is_read_only:
                        # count values may change during the test
                        continue
                    self.assertEqual(getattr(f, field.name), old_values[field.name], (f, field.name))

        # Test refresh of root
        orig_name = self.account.root.name
        self.account.root.name = 'xxx'
        self.account.root.refresh()
        self.assertEqual(self.account.root.name, orig_name)

        folder = Folder()
        with self.assertRaises(ValueError):
            folder.refresh()  # Must have root folder
        folder.root = self.account.root
        with self.assertRaises(ValueError):
            folder.refresh()  # Must have an id

    def test_parent(self):
        self.assertEqual(
            self.account.calendar.parent.name,
            'Top of Information Store'
        )
        self.assertEqual(
            self.account.calendar.parent.parent.name,
            'root'
        )

    def test_children(self):
        self.assertIn(
            'Top of Information Store',
            [c.name for c in self.account.root.children]
        )

    def test_parts(self):
        self.assertEqual(
            [p.name for p in self.account.calendar.parts],
            ['root', 'Top of Information Store', self.account.calendar.name]
        )

    def test_absolute(self):
        self.assertEqual(
            self.account.calendar.absolute,
            '/root/Top of Information Store/' + self.account.calendar.name
        )

    def test_walk(self):
        self.assertGreaterEqual(len(list(self.account.root.walk())), 20)
        self.assertGreaterEqual(len(list(self.account.contacts.walk())), 2)

    def test_tree(self):
        self.assertTrue(self.account.root.tree().startswith('root'))

    def test_glob(self):
        self.assertGreaterEqual(len(list(self.account.root.glob('*'))), 5)
        self.assertEqual(len(list(self.account.contacts.glob('GAL*'))), 1)
        self.assertEqual(len(list(self.account.contacts.glob('gal*'))), 1)  # Test case-insensitivity
        self.assertGreaterEqual(len(list(self.account.contacts.glob('/'))), 5)
        self.assertGreaterEqual(len(list(self.account.contacts.glob('../*'))), 5)
        self.assertEqual(len(list(self.account.root.glob('**/%s' % self.account.contacts.name))), 1)
        self.assertEqual(len(list(self.account.root.glob('Top of*/%s' % self.account.contacts.name))), 1)

    def test_collection_filtering(self):
        self.assertGreaterEqual(self.account.root.tois.children.all().count(), 0)
        self.assertGreaterEqual(self.account.root.tois.walk().all().count(), 0)
        self.assertGreaterEqual(self.account.root.tois.glob('*').all().count(), 0)

    def test_empty_collections(self):
        self.assertEqual(self.account.trash.children.all().count(), 0)
        self.assertEqual(self.account.trash.walk().all().count(), 0)
        self.assertEqual(self.account.trash.glob('XXX').all().count(), 0)
        self.assertEqual(list(self.account.trash.glob('XXX').get_folders()), [])
        self.assertEqual(list(self.account.trash.glob('XXX').find_folders()), [])

    def test_div_navigation(self):
        self.assertEqual(
            (self.account.root / 'Top of Information Store' / self.account.calendar.name).id,
            self.account.calendar.id
        )
        self.assertEqual(
            (self.account.root / 'Top of Information Store' / '..').id,
            self.account.root.id
        )
        self.assertEqual(
            (self.account.root / '.').id,
            self.account.root.id
        )

    def test_double_div_navigation(self):
        self.account.root.clear_cache()  # Clear the cache

        # Test normal navigation
        self.assertEqual(
            (self.account.root // 'Top of Information Store' // self.account.calendar.name).id,
            self.account.calendar.id
        )
        self.assertIsNone(self.account.root._subfolders)

        # Test parent ('..') syntax. Should not work
        with self.assertRaises(ValueError) as e:
            _ = self.account.root // 'Top of Information Store' // '..'
        self.assertEqual(e.exception.args[0], 'Cannot get parent without a folder cache')
        self.assertIsNone(self.account.root._subfolders)

        # Test self ('.') syntax
        self.assertEqual(
            (self.account.root // '.').id,
            self.account.root.id
        )
        self.assertIsNone(self.account.root._subfolders)

    def test_extended_properties(self):
        # Test extended properties on folders and folder roots. This extended prop gets the size (in bytes) of a folder
        class FolderSize(ExtendedProperty):
            property_tag = 0x0e08
            property_type = 'Integer'

        try:
            Folder.register('size', FolderSize)
            self.account.inbox.refresh()
            self.assertGreater(self.account.inbox.size, 0)
        finally:
            Folder.deregister('size')

        try:
            RootOfHierarchy.register('size', FolderSize)
            self.account.root.refresh()
            self.assertGreater(self.account.root.size, 0)
        finally:
            RootOfHierarchy.deregister('size')

        # Register and deregister is only allowed on Folder and RootOfHierarchy classes
        with self.assertRaises(TypeError):
            self.account.calendar.register(FolderSize)
        with self.assertRaises(TypeError):
            self.account.calendar.deregister(FolderSize)
        with self.assertRaises(TypeError):
            self.account.root.register(FolderSize)
        with self.assertRaises(TypeError):
            self.account.root.deregister(FolderSize)

    def test_create_update_empty_delete(self):
        f = Messages(parent=self.account.inbox, name=get_random_string(16))
        f.save()
        self.assertIsNotNone(f.id)
        self.assertIsNotNone(f.changekey)

        new_name = get_random_string(16)
        f.name = new_name
        f.save()
        f.refresh()
        self.assertEqual(f.name, new_name)

        with self.assertRaises(ErrorObjectTypeChanged):
            # FolderClass may not be changed
            f.folder_class = get_random_string(16)
            f.save(update_fields=['folder_class'])

        # Create a subfolder
        Messages(parent=f, name=get_random_string(16)).save()
        self.assertEqual(len(list(f.children)), 1)
        f.empty()
        self.assertEqual(len(list(f.children)), 1)
        f.empty(delete_sub_folders=True)
        self.assertEqual(len(list(f.children)), 0)

        # Create a subfolder again, and delete it by wiping
        Messages(parent=f, name=get_random_string(16)).save()
        self.assertEqual(len(list(f.children)), 1)
        f.wipe()
        self.assertEqual(len(list(f.children)), 0)

        f.delete()
        with self.assertRaises(ValueError):
            # No longer has an ID
            f.refresh()

        # Delete all subfolders of inbox
        for c in self.account.inbox.children:
            c.delete()

        with self.assertRaises(ErrorDeleteDistinguishedFolder):
            self.account.inbox.delete()

    def test_move(self):
        f1 = Folder(parent=self.account.inbox, name=get_random_string(16)).save()
        f2 = Folder(parent=self.account.inbox, name=get_random_string(16)).save()

        f1_id, f1_changekey, f1_parent = f1.id, f1.changekey, f1.parent
        f1.move(f2)
        self.assertEqual(f1.id, f1_id)
        self.assertNotEqual(f1.changekey, f1_changekey)
        self.assertEqual(f1.parent, f2)
        self.assertNotEqual(f1.changekey, f1_parent)

        f1_id, f1_changekey, f1_parent = f1.id, f1.changekey, f1.parent
        f1.refresh()
        self.assertEqual(f1.id, f1_id)
        self.assertEqual(f1.parent, f2)
        self.assertNotEqual(f1.changekey, f1_parent)

        f1.delete()
        f2.delete()

    def test_generic_folder(self):
        f = Folder(parent=self.account.inbox, name=get_random_string(16))
        f.save()
        f.name = get_random_string(16)
        f.save()
        f.delete()

    def test_folder_query_set(self):
        # Create a folder hierarchy and test a folder queryset
        #
        # -f0
        #  - f1
        #  - f2
        #    - f21
        #    - f22
        f0 = Folder(parent=self.account.inbox, name=get_random_string(16)).save()
        f1 = Folder(parent=f0, name=get_random_string(16)).save()
        f2 = Folder(parent=f0, name=get_random_string(16)).save()
        f21 = Folder(parent=f2, name=get_random_string(16)).save()
        f22 = Folder(parent=f2, name=get_random_string(16)).save()
        folder_qs = SingleFolderQuerySet(account=self.account, folder=f0)
        try:
            # Test all()
            self.assertSetEqual(
                {f.name for f in folder_qs.all()},
                {f.name for f in (f1, f2, f21, f22)}
            )

            # Test only()
            self.assertSetEqual(
                {f.name for f in folder_qs.only('name').all()},
                {f.name for f in (f1, f2, f21, f22)}
            )
            self.assertSetEqual(
                {f.child_folder_count for f in folder_qs.only('name').all()},
                {None}
            )
            # Test depth()
            self.assertSetEqual(
                {f.name for f in folder_qs.depth(SHALLOW).all()},
                {f.name for f in (f1, f2)}
            )

            # Test filter()
            self.assertSetEqual(
                {f.name for f in folder_qs.filter(name=f1.name)},
                {f.name for f in (f1,)}
            )
            self.assertSetEqual(
                {f.name for f in folder_qs.filter(name__in=[f1.name, f2.name])},
                {f.name for f in (f1, f2)}
            )

            # Test get()
            self.assertEqual(folder_qs.get(id=f2.id).name, f2.name)
            self.assertEqual(folder_qs.get(id=f2.id, changekey=f2.changekey).name, f2.name)
            self.assertEqual(
                folder_qs.get(name=f2.name).child_folder_count,
                2
            )
            self.assertEqual(
                folder_qs.filter(name=f2.name).get().child_folder_count,
                2
            )
            self.assertEqual(
                folder_qs.only('name').get(name=f2.name).name,
                f2.name
            )
            self.assertEqual(
                folder_qs.only('name').get(name=f2.name).child_folder_count,
                None
            )
            with self.assertRaises(DoesNotExist):
                folder_qs.get(name=get_random_string(16))
            with self.assertRaises(MultipleObjectsReturned):
                folder_qs.get()
        finally:
            f0.wipe()
            f0.delete()

    def test_folder_query_set_failures(self):
        with self.assertRaises(ValueError):
            FolderQuerySet('XXX')
        fld_qs = SingleFolderQuerySet(account=self.account, folder=self.account.inbox)
        with self.assertRaises(InvalidField):
            fld_qs.only('XXX')
        with self.assertRaises(InvalidField):
            list(fld_qs.filter(XXX='XXX'))

    def test_user_configuration(self):
        """Test that we can do CRUD operations on user configuration data."""
        # Create a test folder that we delete afterwards
        f = Messages(parent=self.account.inbox, name=get_random_string(16)).save()
        # The name must be fewer than 237 characters, can contain only the characters "A-Z", "a-z", "0-9", and ".",
        # and must not start with "IPM.Configuration"
        name = get_random_string(16, spaces=False, special=False)

        # Should not exist yet
        with self.assertRaises(ErrorItemNotFound):
            f.get_user_configuration(name=name)

        # Create a config
        dictionary = {
            get_random_bool(): get_random_str_tuple(10, 2),
            get_random_int(): get_random_bool(),
            get_random_byte(): get_random_int(),
            get_random_bytes(16): get_random_byte(),
            get_random_string(8): get_random_bytes(16),
            get_random_datetime(tz=self.account.default_timezone): get_random_string(8),
            get_random_str_tuple(4, 4): get_random_datetime(tz=self.account.default_timezone),
        }
        xml_data = b'<foo>%s</foo>' % get_random_string(16).encode('utf-8')
        binary_data = get_random_bytes(100)
        f.create_user_configuration(name=name, dictionary=dictionary, xml_data=xml_data, binary_data=binary_data)

        # Fetch and compare values
        config = f.get_user_configuration(name=name)
        self.assertEqual(config.dictionary, dictionary)
        self.assertEqual(config.xml_data, xml_data)
        self.assertEqual(config.binary_data, binary_data)

        # Cannot create one more with the same name
        with self.assertRaises(ErrorItemSave):
            f.create_user_configuration(name=name)

        # Does not exist on a different folder
        with self.assertRaises(ErrorItemNotFound):
            self.account.inbox.get_user_configuration(name=name)

        # Update the config
        f.update_user_configuration(
            name=name, dictionary={'bar': 'foo', 456: 'a', 'b': True}, xml_data=b'<foo>baz</foo>', binary_data=b'YYY'
        )

        # Fetch again and compare values
        config = f.get_user_configuration(name=name)
        self.assertEqual(config.dictionary, {'bar': 'foo', 456: 'a', 'b': True})
        self.assertEqual(config.xml_data, b'<foo>baz</foo>')
        self.assertEqual(config.binary_data, b'YYY')

        # Delete the config
        f.delete_user_configuration(name=name)

        # We already deleted this config
        with self.assertRaises(ErrorItemNotFound):
            f.get_user_configuration(name=name)
        f.delete()

    def test_permissionset_effectiverights_parsing(self):
        # Test static XML since server may not have any permission sets or effective rights
        xml = b'''\
<?xml version="1.0" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:GetFolderResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
                <m:GetFolderResponseMessage ResponseClass="Success">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:Folders>
                        <t:CalendarFolder>
                            <t:FolderId Id="XXX" ChangeKey="YYY"/>
                            <t:ParentFolderId Id="ZZZ" ChangeKey="WWW"/>
                            <t:FolderClass>IPF.Appointment</t:FolderClass>
                            <t:DisplayName>My Calendar</t:DisplayName>
                            <t:TotalCount>1</t:TotalCount>
                            <t:ChildFolderCount>2</t:ChildFolderCount>
                            <t:EffectiveRights>
                                <t:CreateAssociated>true</t:CreateAssociated>
                                <t:CreateContents>true</t:CreateContents>
                                <t:CreateHierarchy>true</t:CreateHierarchy>
                                <t:Delete>true</t:Delete>
                                <t:Modify>true</t:Modify>
                                <t:Read>true</t:Read>
                                <t:ViewPrivateItems>false</t:ViewPrivateItems>
                            </t:EffectiveRights>
                            <t:PermissionSet>
                                <t:CalendarPermissions>
                                    <t:CalendarPermission>
                                        <t:UserId>
                                            <t:SID>SID1</t:SID>
                                            <t:PrimarySmtpAddress>user1@example.com</t:PrimarySmtpAddress>
                                            <t:DisplayName>User 1</t:DisplayName>
                                        </t:UserId>
                                        <t:CanCreateItems>false</t:CanCreateItems>
                                        <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                                        <t:IsFolderOwner>false</t:IsFolderOwner>
                                        <t:IsFolderVisible>true</t:IsFolderVisible>
                                        <t:IsFolderContact>false</t:IsFolderContact>
                                        <t:EditItems>None</t:EditItems>
                                        <t:DeleteItems>None</t:DeleteItems>
                                        <t:ReadItems>FullDetails</t:ReadItems>
                                        <t:CalendarPermissionLevel>Reviewer</t:CalendarPermissionLevel>
                                    </t:CalendarPermission>
                                    <t:CalendarPermission>
                                        <t:UserId>
                                            <t:SID>SID2</t:SID>
                                            <t:PrimarySmtpAddress>user2@example.com</t:PrimarySmtpAddress>
                                            <t:DisplayName>User 2</t:DisplayName>
                                        </t:UserId>
                                        <t:CanCreateItems>true</t:CanCreateItems>
                                        <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                                        <t:IsFolderOwner>false</t:IsFolderOwner>
                                        <t:IsFolderVisible>true</t:IsFolderVisible>
                                        <t:IsFolderContact>false</t:IsFolderContact>
                                        <t:EditItems>All</t:EditItems>
                                        <t:DeleteItems>All</t:DeleteItems>
                                        <t:ReadItems>FullDetails</t:ReadItems>
                                        <t:CalendarPermissionLevel>Editor</t:CalendarPermissionLevel>
                                    </t:CalendarPermission>
                                </t:CalendarPermissions>
                            </t:PermissionSet>
                        </t:CalendarFolder>
                    </m:Folders>
                </m:GetFolderResponseMessage>
            </m:ResponseMessages>
        </m:GetFolderResponse>
    </s:Body>
</s:Envelope>'''
        ws = GetFolder(account=self.account)
        ws.folders = [self.account.calendar]
        res = list(ws.parse(xml))
        self.assertEqual(len(res), 1)
        fld = res[0]
        self.assertEqual(
            fld.effective_rights,
            EffectiveRights(
                create_associated=True, create_contents=True, create_hierarchy=True, delete=True, modify=True,
                read=True, view_private_items=False
            )
        )
        self.assertEqual(
            fld.permission_set,
            PermissionSet(
                permissions=None,
                calendar_permissions=[
                    CalendarPermission(
                        can_create_items=False, can_create_subfolders=False, is_folder_owner=False,
                        is_folder_visible=True, is_folder_contact=False, edit_items='None',
                        delete_items='None', read_items='FullDetails', user_id=UserId(
                            sid='SID1', primary_smtp_address='user1@example.com', display_name='User 1',
                            distinguished_user=None, external_user_identity=None
                        ), calendar_permission_level='Reviewer'
                    ),
                    CalendarPermission(
                        can_create_items=True, can_create_subfolders=False, is_folder_owner=False,
                        is_folder_visible=True, is_folder_contact=False, edit_items='All', delete_items='All',
                        read_items='FullDetails', user_id=UserId(
                            sid='SID2', primary_smtp_address='user2@example.com', display_name='User 2',
                            distinguished_user=None, external_user_identity=None
                        ), calendar_permission_level='Editor'
                    )
                ], unknown_entries=None
            ),
        )
