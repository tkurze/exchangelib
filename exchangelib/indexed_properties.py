import logging

from .fields import EmailSubField, LabelField, SubField, NamedSubField, Choice
from .properties import EWSElement, EWSMeta

log = logging.getLogger(__name__)


class IndexedElement(EWSElement, metaclass=EWSMeta):
    """Base class for all classes that implement an indexed element."""

    LABELS = set()


class SingleFieldIndexedElement(IndexedElement, metaclass=EWSMeta):
    """Base class for all classes that implement an indexed element with a single field."""

    @classmethod
    def value_field(cls, version=None):
        fields = cls.supported_fields(version=version)
        if len(fields) != 1:
            raise ValueError('This class must have only one field (found %s)' % (fields,))
        return fields[0]


class EmailAddress(SingleFieldIndexedElement):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/entry-emailaddress"""

    ELEMENT_NAME = 'Entry'
    LABEL_CHOICES = ('EmailAddress1', 'EmailAddress2', 'EmailAddress3')

    label = LabelField(field_uri='Key', choices={Choice(c) for c in LABEL_CHOICES}, default=LABEL_CHOICES[0])
    email = EmailSubField()


class PhoneNumber(SingleFieldIndexedElement):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/entry-phonenumber"""

    ELEMENT_NAME = 'Entry'

    label = LabelField(field_uri='Key', choices={
        Choice('AssistantPhone'), Choice('BusinessFax'), Choice('BusinessPhone'), Choice('BusinessPhone2'),
        Choice('Callback'), Choice('CarPhone'), Choice('CompanyMainPhone'), Choice('HomeFax'), Choice('HomePhone'),
        Choice('HomePhone2'), Choice('Isdn'), Choice('MobilePhone'), Choice('OtherFax'), Choice('OtherTelephone'),
        Choice('Pager'), Choice('PrimaryPhone'), Choice('RadioPhone'), Choice('Telex'), Choice('TtyTddPhone'),
    }, default='PrimaryPhone')
    phone_number = SubField()


class MultiFieldIndexedElement(IndexedElement, metaclass=EWSMeta):
    """Base class for all classes that implement an indexed element with multiple fields."""


class PhysicalAddress(MultiFieldIndexedElement):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/entry-physicaladdress"""

    ELEMENT_NAME = 'Entry'

    label = LabelField(field_uri='Key', choices={
        Choice('Business'), Choice('Home'), Choice('Other')
    }, default='Business')
    street = NamedSubField(field_uri='Street')  # Street, house number, etc.
    city = NamedSubField(field_uri='City')
    state = NamedSubField(field_uri='State')
    country = NamedSubField(field_uri='CountryOrRegion')
    zipcode = NamedSubField(field_uri='PostalCode')

    def clean(self, version=None):
        if isinstance(self.zipcode, int):
            self.zipcode = str(self.zipcode)
        super().clean(version=version)
