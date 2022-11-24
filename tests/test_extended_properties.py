from exchangelib.extended_properties import ExtendedProperty, Flag
from exchangelib.folders import Inbox
from exchangelib.items import BaseItem, CalendarItem, Message
from exchangelib.properties import Mailbox
from exchangelib.util import to_xml

from .common import get_random_int, get_random_url
from .test_items.test_basics import BaseItemTest


class ExtendedPropertyTest(BaseItemTest):
    TEST_FOLDER = "inbox"
    FOLDER_CLASS = Inbox
    ITEM_CLASS = Message

    def test_register(self):
        # Tests that we can register and de-register custom extended properties
        class TestProp(ExtendedProperty):
            property_set_id = "deadbeaf-cafe-cafe-cafe-deadbeefcafe"
            property_name = "Test Property"
            property_type = "Integer"

        attr_name = "dead_beef"

        # Before register
        self.assertNotIn(attr_name, {f.name for f in self.ITEM_CLASS.supported_fields(self.account.version)})
        with self.assertRaises(ValueError):
            self.ITEM_CLASS.deregister(attr_name)  # Not registered yet
        with self.assertRaises(ValueError):
            self.ITEM_CLASS.deregister("subject")  # Not an extended property

        # Test that we can clean an item before and after registry
        item = self.ITEM_CLASS()
        item.clean(version=self.account.version)
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=TestProp)
        item.clean(version=self.account.version)
        try:
            # After register
            self.assertEqual(TestProp.python_type(), int)
            self.assertIn(attr_name, {f.name for f in self.ITEM_CLASS.supported_fields(self.account.version)})

            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.dead_beef
            self.assertIsInstance(prop_val, int)
            item.save()
            item.refresh()
            self.assertEqual(prop_val, item.dead_beef)
            new_prop_val = get_random_int(0, 256)
            item.dead_beef = new_prop_val
            item.save()
            item.refresh()
            self.assertEqual(new_prop_val, item.dead_beef)

            # Test deregister
            with self.assertRaises(ValueError) as e:
                self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=TestProp)
            self.assertEqual(e.exception.args[0], "'attr_name' 'dead_beef' is already registered")
            with self.assertRaises(TypeError) as e:
                self.ITEM_CLASS.register(attr_name="XXX", attr_cls=Mailbox)
            self.assertEqual(
                e.exception.args[0],
                "'attr_cls' <class 'exchangelib.properties.Mailbox'> must be a subclass of type "
                "<class 'exchangelib.extended_properties.ExtendedProperty'>",
            )
            with self.assertRaises(ValueError) as e:
                BaseItem.register(attr_name=attr_name, attr_cls=Mailbox)
                self.assertEqual(
                    e.exception.args[0],
                    "Class <class 'exchangelib.items.base.BaseItem'> is missing INSERT_AFTER_FIELD value",
                )
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)
        self.assertNotIn(attr_name, {f.name for f in self.ITEM_CLASS.supported_fields(self.account.version)})

    def test_extended_property_arraytype(self):
        # Tests array type extended properties
        class TestArayProp(ExtendedProperty):
            property_set_id = "deadcafe-beef-beef-beef-deadcafebeef"
            property_name = "Test Array Property"
            property_type = "IntegerArray"

        attr_name = "dead_beef_array"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=TestArayProp)
        try:
            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.dead_beef_array
            self.assertIsInstance(prop_val, list)
            item.save()
            item.refresh()
            self.assertEqual(prop_val, item.dead_beef_array)
            new_prop_val = self.random_val(self.ITEM_CLASS.get_field_by_fieldname(attr_name))
            item.dead_beef_array = new_prop_val
            item.save()
            item.refresh()
            self.assertEqual(new_prop_val, item.dead_beef_array)
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_extended_property_with_tag(self):
        attr_name = "my_flag"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=Flag)
        try:
            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.my_flag
            self.assertIsInstance(prop_val, int)
            item.save()
            item.refresh()
            self.assertEqual(prop_val, item.my_flag)
            new_prop_val = self.random_val(self.ITEM_CLASS.get_field_by_fieldname(attr_name))
            item.my_flag = new_prop_val
            item.save()
            item.refresh()
            self.assertEqual(new_prop_val, item.my_flag)
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_extended_property_with_invalid_tag(self):
        class InvalidProp(ExtendedProperty):
            property_tag = "0x8000"
            property_type = "Integer"

        with self.assertRaises(ValueError):
            InvalidProp("Foo").clean()  # property_tag is in protected range

    def test_extended_property_with_string_tag(self):
        attr_name = "my_flag"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=Flag)
        try:
            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.my_flag
            self.assertIsInstance(prop_val, int)
            item.save()
            item.refresh()
            self.assertEqual(prop_val, item.my_flag)
            new_prop_val = self.random_val(self.ITEM_CLASS.get_field_by_fieldname(attr_name))
            item.my_flag = new_prop_val
            item.save()
            item.refresh()
            self.assertEqual(new_prop_val, item.my_flag)
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_extended_distinguished_property(self):
        if self.ITEM_CLASS == CalendarItem:
            # MyMeeting is an extended prop version of the 'CalendarItem.uid' field. They don't work together.
            self.skipTest("This extendedproperty doesn't work on CalendarItems")

        class MyMeeting(ExtendedProperty):
            distinguished_property_set_id = "Meeting"
            property_type = "Binary"
            property_id = 3

        attr_name = "my_meeting"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=MyMeeting)
        try:
            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.my_meeting
            self.assertIsInstance(prop_val, bytes)
            item.save()
            item = self.get_item_by_id((item.id, item.changekey))
            self.assertEqual(prop_val, item.my_meeting, (prop_val, item.my_meeting))
            new_prop_val = self.random_val(self.ITEM_CLASS.get_field_by_fieldname(attr_name))
            item.my_meeting = new_prop_val
            item.save()
            item = self.get_item_by_id((item.id, item.changekey))
            self.assertEqual(new_prop_val, item.my_meeting)
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_extended_property_binary_array(self):
        class MyMeetingArray(ExtendedProperty):
            property_set_id = "00062004-0000-0000-C000-000000000046"
            property_type = "BinaryArray"
            property_id = 32852

        attr_name = "my_meeting_array"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=MyMeetingArray)

        try:
            # Test item creation, refresh, and update
            item = self.get_test_item(folder=self.test_folder)
            prop_val = item.my_meeting_array
            self.assertIsInstance(prop_val, list)
            item.save()
            item = self.get_item_by_id((item.id, item.changekey))
            self.assertEqual(prop_val, item.my_meeting_array)
            new_prop_val = self.random_val(self.ITEM_CLASS.get_field_by_fieldname(attr_name))
            item.my_meeting_array = new_prop_val
            item.save()
            item = self.get_item_by_id((item.id, item.changekey))
            self.assertEqual(new_prop_val, item.my_meeting_array)
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)

    def test_extended_property_validation(self):
        # Must not have property_set_id or property_tag
        class TestProp1(ExtendedProperty):
            distinguished_property_set_id = "XXX"
            property_set_id = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp1.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'distinguished_property_set_id' is set, 'property_set_id' and 'property_tag' must be None",
        )

        # Must have property_id or property_name
        class TestProp2(ExtendedProperty):
            distinguished_property_set_id = "XXX"

        with self.assertRaises(ValueError) as e:
            TestProp2.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'distinguished_property_set_id' is set, 'property_id' or 'property_name' must also be set",
        )

        # distinguished_property_set_id must have a valid value
        class TestProp3(ExtendedProperty):
            distinguished_property_set_id = "XXX"
            property_id = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp3.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            f"'distinguished_property_set_id' 'XXX' must be one of {sorted(ExtendedProperty.DISTINGUISHED_SETS)}",
        )

        # Must not have distinguished_property_set_id or property_tag
        class TestProp4(ExtendedProperty):
            property_set_id = "XXX"
            property_tag = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp4.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'property_set_id' is set, 'distinguished_property_set_id' and 'property_tag' must be None",
        )

        # Must have property_id or property_name
        class TestProp5(ExtendedProperty):
            property_set_id = "XXX"

        with self.assertRaises(ValueError) as e:
            TestProp5.validate_cls()
        self.assertEqual(
            e.exception.args[0], "When 'property_set_id' is set, 'property_id' or 'property_name' must also be set"
        )

        # property_tag is only compatible with property_type
        class TestProp6(ExtendedProperty):
            property_tag = "XXX"
            property_set_id = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp6.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'property_set_id' is set, 'distinguished_property_set_id' and 'property_tag' must be None",
        )

        # property_tag must be an integer or string that can be converted to int
        class TestProp7(ExtendedProperty):
            property_tag = "XXX"

        with self.assertRaises(ValueError) as e:
            TestProp7.validate_cls()
        self.assertEqual(e.exception.args[0], "invalid literal for int() with base 16: 'XXX'")

        # property_tag must not be in the reserved range
        class TestProp8(ExtendedProperty):
            property_tag = 0x8001

        with self.assertRaises(ValueError) as e:
            TestProp8.validate_cls()
        self.assertEqual(e.exception.args[0], "'property_tag' value '0x8001' is reserved for custom properties")

        # Must not have property_id or property_tag
        class TestProp9(ExtendedProperty):
            property_name = "XXX"
            property_id = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp9.validate_cls()
        self.assertEqual(
            e.exception.args[0], "When 'property_name' is set, 'property_id' and 'property_tag' must be None"
        )

        # Must have distinguished_property_set_id or property_set_id
        class TestProp10(ExtendedProperty):
            property_name = "XXX"

        with self.assertRaises(ValueError) as e:
            TestProp10.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'property_name' is set, 'distinguished_property_set_id' or 'property_set_id' must also be set",
        )

        # Must not have property_name or property_tag
        class TestProp11(ExtendedProperty):
            property_id = "XXX"
            property_name = "YYY"

        with self.assertRaises(ValueError) as e:
            TestProp11.validate_cls()  # This actually hits the check on property_name values
        self.assertEqual(
            e.exception.args[0], "When 'property_name' is set, 'property_id' and 'property_tag' must be None"
        )

        # Must have distinguished_property_set_id or property_set_id
        class TestProp12(ExtendedProperty):
            property_id = "XXX"

        with self.assertRaises(ValueError) as e:
            TestProp12.validate_cls()
        self.assertEqual(
            e.exception.args[0],
            "When 'property_id' is set, 'distinguished_property_set_id' or 'property_set_id' must also be set",
        )

        # property_type must be a valid value
        class TestProp13(ExtendedProperty):
            property_id = "XXX"
            property_set_id = "YYY"
            property_type = "ZZZ"

        with self.assertRaises(ValueError) as e:
            TestProp13.validate_cls()
        self.assertEqual(
            e.exception.args[0], f"'property_type' 'ZZZ' must be one of {sorted(ExtendedProperty.PROPERTY_TYPES)}"
        )

        # property_tag and property_id are mutually exclusive
        class TestProp14(ExtendedProperty):
            property_tag = "XXX"
            property_id = "YYY"

        with self.assertRaises(ValueError) as e:
            # We cannot reach this exception directly with validate_cls()
            TestProp14._validate_property_tag()
        self.assertEqual(e.exception.args[0], "When 'property_tag' is set, only 'property_type' must be set")
        with self.assertRaises(ValueError) as e:
            # We cannot reach this exception directly with validate_cls()
            TestProp14._validate_property_id()
        self.assertEqual(
            e.exception.args[0], "When 'property_id' is set, 'property_name' and 'property_tag' must be None"
        )

    def test_multiple_extended_properties(self):
        class ExternalSharingUrl(ExtendedProperty):
            property_set_id = "F52A8693-C34D-4980-9E20-9D4C1EABB6A7"
            property_name = "ExternalSharingUrl"
            property_type = "String"

        class ExternalSharingFolderId(ExtendedProperty):
            property_set_id = "F52A8693-C34D-4980-9E20-9D4C1EABB6A7"
            property_name = "ExternalSharingLocalFolderId"
            property_type = "Binary"

        try:
            self.ITEM_CLASS.register("sharing_url", ExternalSharingUrl)
            self.ITEM_CLASS.register("sharing_folder_id", ExternalSharingFolderId)

            url, folder_id = get_random_url(), self.test_folder.id.encode("utf-8")
            m = self.get_test_item()
            m.sharing_url, m.sharing_folder_id = url, folder_id
            m.save()

            m = self.test_folder.get(sharing_url=url)
            self.assertEqual(m.sharing_url, url)
            self.assertEqual(m.sharing_folder_id, folder_id)
        finally:
            self.ITEM_CLASS.deregister("sharing_url")
            self.ITEM_CLASS.deregister("sharing_folder_id")

    def test_via_queryset(self):
        class TestProp(ExtendedProperty):
            property_set_id = "deadbeaf-cafe-cafe-cafe-deadbeefcafe"
            property_name = "Test Property"
            property_type = "Integer"

        class TestArayProp(ExtendedProperty):
            property_set_id = "deadcafe-beef-beef-beef-deadcafebeef"
            property_name = "Test Array Property"
            property_type = "IntegerArray"

        attr_name = "dead_beef"
        array_attr_name = "dead_beef_array"
        self.ITEM_CLASS.register(attr_name=attr_name, attr_cls=TestProp)
        self.ITEM_CLASS.register(attr_name=array_attr_name, attr_cls=TestArayProp)
        try:
            item = self.get_test_item(folder=self.test_folder).save()
            self.assertEqual(self.test_folder.filter(**{attr_name: getattr(item, attr_name)}).count(), 1)
            self.assertEqual(
                # Does not work in O365
                # self.test_folder.filter(**{f"{array_attr_name}__contains": getattr(item, array_attr_name)}).count(), 1
                self.test_folder.filter(**{f"{array_attr_name}__in": getattr(item, array_attr_name)}).count(),
                1,
            )
        finally:
            self.ITEM_CLASS.deregister(attr_name=attr_name)
            self.ITEM_CLASS.deregister(attr_name=array_attr_name)

    def test_from_xml(self):
        # Test that empty and no-op XML Value elements for string props both return empty strings
        class TestProp(ExtendedProperty):
            property_set_id = "deadbeaf-cafe-cafe-cafe-deadbeefcafe"
            property_name = "Test Property"
            property_type = "String"

        elem = to_xml(
            b"""\
<ExtendedProperty xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
    <ExtendedFieldURI/>
    <Value>XXX</Value>
</ExtendedProperty>"""
        )
        self.assertEqual(TestProp.from_xml(elem, account=None), "XXX")
        elem = to_xml(
            b"""\
<ExtendedProperty xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
    <ExtendedFieldURI/>
    <Value></Value>
</ExtendedProperty>"""
        )
        self.assertEqual(TestProp.from_xml(elem, account=None), "")
        elem = to_xml(
            b"""\
<ExtendedProperty xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
    <ExtendedFieldURI/>
    <Value/>
</ExtendedProperty>"""
        )
        self.assertEqual(TestProp.from_xml(elem, account=None), "")
