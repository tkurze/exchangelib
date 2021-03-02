import datetime
import logging

from .item import Item
from ..fields import BooleanField, Base64Field, TextField, ChoiceField, URIField, DateTimeBackedDateField, \
    PhoneNumberField, EmailAddressesField, PhysicalAddressField, Choice, MemberListField, CharField, TextListField, \
    EmailAddressField, IdElementField, EWSElementField, DateTimeField, EWSElementListField, \
    BodyContentAttributedValueField, StringAttributedValueField, PhoneNumberAttributedValueField, \
    PersonaPhoneNumberField, EmailAddressAttributedValueField, PostalAddressAttributedValueField
from ..properties import PersonaId, IdChangeKeyMixIn, Fields, CompleteName, Attribution, EmailAddress, Address, \
    FolderId
from ..util import TNS
from ..version import EXCHANGE_2010, EXCHANGE_2010_SP2, EXCHANGE_2013

log = logging.getLogger(__name__)


class Contact(Item):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/contact"""
    ELEMENT_NAME = 'Contact'
    LOCAL_FIELDS = Fields(
        TextField('file_as', field_uri='contacts:FileAs'),
        ChoiceField('file_as_mapping', field_uri='contacts:FileAsMapping', choices={
            Choice('None'), Choice('LastCommaFirst'), Choice('FirstSpaceLast'), Choice('Company'),
            Choice('LastCommaFirstCompany'), Choice('CompanyLastFirst'), Choice('LastFirst'),
            Choice('LastFirstCompany'), Choice('CompanyLastCommaFirst'), Choice('LastFirstSuffix'),
            Choice('LastSpaceFirstCompany'), Choice('CompanyLastSpaceFirst'), Choice('LastSpaceFirst'),
            Choice('DisplayName'), Choice('FirstName'), Choice('LastFirstMiddleSuffix'), Choice('LastName'),
            Choice('Empty'),
        }),
        TextField('display_name', field_uri='contacts:DisplayName', is_required=True),
        CharField('given_name', field_uri='contacts:GivenName'),
        TextField('initials', field_uri='contacts:Initials'),
        CharField('middle_name', field_uri='contacts:MiddleName'),
        TextField('nickname', field_uri='contacts:Nickname'),
        EWSElementField('complete_name', field_uri='contacts:CompleteName', value_cls=CompleteName, is_read_only=True),
        TextField('company_name', field_uri='contacts:CompanyName'),
        EmailAddressesField('email_addresses', field_uri='contacts:EmailAddress'),
        PhysicalAddressField('physical_addresses', field_uri='contacts:PhysicalAddress'),
        PhoneNumberField('phone_numbers', field_uri='contacts:PhoneNumber'),
        TextField('assistant_name', field_uri='contacts:AssistantName'),
        DateTimeBackedDateField('birthday', field_uri='contacts:Birthday', default_time=datetime.time(11, 59)),
        URIField('business_homepage', field_uri='contacts:BusinessHomePage'),
        TextListField('children', field_uri='contacts:Children'),
        TextListField('companies', field_uri='contacts:Companies', is_searchable=False),
        ChoiceField('contact_source', field_uri='contacts:ContactSource', choices={
            Choice('Store'), Choice('ActiveDirectory')
        }, is_read_only=True),
        TextField('department', field_uri='contacts:Department'),
        TextField('generation', field_uri='contacts:Generation'),
        CharField('im_addresses', field_uri='contacts:ImAddresses', is_read_only=True),
        TextField('job_title', field_uri='contacts:JobTitle'),
        TextField('manager', field_uri='contacts:Manager'),
        TextField('mileage', field_uri='contacts:Mileage'),
        TextField('office', field_uri='contacts:OfficeLocation'),
        ChoiceField('postal_address_index', field_uri='contacts:PostalAddressIndex', choices={
            Choice('Business'), Choice('Home'), Choice('Other'), Choice('None')
        }, default='None', is_required_after_save=True),
        TextField('profession', field_uri='contacts:Profession'),
        TextField('spouse_name', field_uri='contacts:SpouseName'),
        CharField('surname', field_uri='contacts:Surname'),
        DateTimeBackedDateField('wedding_anniversary', field_uri='contacts:WeddingAnniversary',
                                default_time=datetime.time(11, 59)),
        BooleanField('has_picture', field_uri='contacts:HasPicture', supported_from=EXCHANGE_2010, is_read_only=True),
        TextField('phonetic_full_name', field_uri='contacts:PhoneticFullName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        TextField('phonetic_first_name', field_uri='contacts:PhoneticFirstName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        TextField('phonetic_last_name', field_uri='contacts:PhoneticLastName', supported_from=EXCHANGE_2013,
                  is_read_only=True),
        EmailAddressField('email_alias', field_uri='contacts:Alias', is_read_only=True),
        # 'notes' is documented in MSDN but apparently unused. Writing to it raises ErrorInvalidPropertyRequest. OWA
        # put entries into the 'notes' form field into the 'body' field.
        CharField('notes', field_uri='contacts:Notes', supported_from=EXCHANGE_2013, is_read_only=True),
        # 'photo' is documented in MSDN but apparently unused. Writing to it raises ErrorInvalidPropertyRequest. OWA
        # adds photos as FileAttachments on the contact item (with 'is_contact_photo=True'), which automatically flips
        # the 'has_picture' field.
        Base64Field('photo', field_uri='contacts:Photo', is_read_only=True),
        Base64Field('user_smime_certificate', field_uri='contacts:UserSMIMECertificate', is_read_only=True,
                    supported_from=EXCHANGE_2010_SP2),
        Base64Field('ms_exchange_certificate', field_uri='contacts:MSExchangeCertificate', is_read_only=True,
                    supported_from=EXCHANGE_2010_SP2),
        TextField('directory_id', field_uri='contacts:DirectoryId', supported_from=EXCHANGE_2013, is_read_only=True),
        CharField('manager_mailbox', field_uri='contacts:ManagerMailbox', supported_from=EXCHANGE_2010_SP2,
                  is_read_only=True),
        CharField('direct_reports', field_uri='contacts:DirectReports', supported_from=EXCHANGE_2010_SP2,
                  is_read_only=True),
    )
    FIELDS = Item.FIELDS + LOCAL_FIELDS

    __slots__ = tuple(f.name for f in LOCAL_FIELDS)


class Persona(IdChangeKeyMixIn):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/persona"""
    ELEMENT_NAME = 'Persona'
    ID_ELEMENT_CLS = PersonaId
    LOCAL_FIELDS = Fields(
        IdElementField('_id', field_uri='persona:PersonaId', value_cls=ID_ELEMENT_CLS, namespace=TNS),
        CharField('persona_type', field_uri='persona:PersonaType'),
        TextField('persona_object_type', field_uri='persona:PersonaObjectStatus'),
        DateTimeField('creation_time', field_uri='persona:CreationTime'),
        BodyContentAttributedValueField('bodies', field_uri='persona:Bodies'),
        TextField('display_name_first_last_sort_key', field_uri='persona:DisplayNameFirstLastSortKey'),
        TextField('display_name_last_first_sort_key', field_uri='persona:DisplayNameLastFirstSortKey'),
        TextField('company_sort_key', field_uri='persona:CompanyNameSortKey'),
        TextField('home_sort_key', field_uri='persona:HomeCitySortKey'),
        TextField('work_city_sort_key', field_uri='persona:WorkCitySortKey'),
        CharField('display_name_first_last_header', field_uri='persona:DisplayNameFirstLastHeader'),
        CharField('display_name_last_first_header', field_uri='persona:DisplayNameLastFirstHeader'),
        TextField('file_as_header', field_uri='persona:FileAsHeader'),
        CharField('display_name', field_uri='persona:DisplayName'),
        CharField('display_name_first_last', field_uri='persona:DisplayNameFirstLast'),
        CharField('display_name_last_first', field_uri='persona:DisplayNameLastFirst'),
        CharField('file_as', field_uri='persona:FileAs'),
        TextField('file_as_id', field_uri='persona:FileAsId'),
        CharField('display_name_prefix', field_uri='persona:DisplayNamePrefix'),
        CharField('given_name', field_uri='persona:GivenName'),
        CharField('middle_name', field_uri='persona:MiddleName'),
        CharField('surname', field_uri='persona:Surname'),
        CharField('generation', field_uri='persona:Generation'),
        TextField('nickname', field_uri='persona:Nickname'),
        TextField('yomi_company_name', field_uri='persona:YomiCompanyName'),
        TextField('yomi_first_name', field_uri='persona:YomiFirstName'),
        TextField('yomi_last_name', field_uri='persona:YomiLastName'),
        CharField('title', field_uri='persona:Title'),
        TextField('department', field_uri='persona:Department'),
        CharField('company_name', field_uri='persona:CompanyName'),
        EWSElementField('email_address', field_uri='persona:EmailAddress', value_cls=EmailAddress),
        EWSElementListField('email_addresses', field_uri='persona:EmailAddresses', value_cls=Address),
        PersonaPhoneNumberField('PhoneNumber', field_uri='persona:PhoneNumber'),
        CharField('im_address', field_uri='persona:ImAddress'),
        CharField('home_city', field_uri='persona:HomeCity'),
        CharField('work_city', field_uri='persona:WorkCity'),
        CharField('relevance_score', field_uri='persona:RelevanceScore'),
        EWSElementListField('folder_ids', field_uri='persona:FolderIds', value_cls=FolderId),
        EWSElementListField('attributions', field_uri='persona:Attributions', value_cls=Attribution),
        StringAttributedValueField('display_names', field_uri='persona:DisplayNames'),
        StringAttributedValueField('file_ases', field_uri='persona:FileAses'),
        StringAttributedValueField('file_as_ids', field_uri='persona:FileAsIds'),
        StringAttributedValueField('display_name_prefixes', field_uri='persona:DisplayNamePrefixes'),
        StringAttributedValueField('given_names', field_uri='persona:GivenNames'),
        StringAttributedValueField('middle_names', field_uri='persona:MiddleNames'),
        StringAttributedValueField('surnames', field_uri='persona:Surnames'),
        StringAttributedValueField('generations', field_uri='persona:Generations'),
        StringAttributedValueField('nicknames', field_uri='persona:Nicknames'),
        StringAttributedValueField('initials', field_uri='persona:Initials'),
        StringAttributedValueField('yomi_company_names', field_uri='persona:YomiCompanyNames'),
        StringAttributedValueField('yomi_first_names', field_uri='persona:YomiFirstNames'),
        StringAttributedValueField('yomi_last_names', field_uri='persona:YomiLastNames'),
        PhoneNumberAttributedValueField('business_phone_numbers', field_uri='persona:BusinessPhoneNumbers'),
        PhoneNumberAttributedValueField('business_phone_numbers2', field_uri='persona:BusinessPhoneNumbers2'),
        PhoneNumberAttributedValueField('home_phones', field_uri='persona:HomePhones'),
        PhoneNumberAttributedValueField('home_phones2', field_uri='persona:HomePhones2'),
        PhoneNumberAttributedValueField('mobile_phones', field_uri='persona:MobilePhones'),
        PhoneNumberAttributedValueField('mobile_phones2', field_uri='persona:MobilePhones2'),
        PhoneNumberAttributedValueField('assistant_phone_numbers', field_uri='persona:AssistantPhoneNumbers'),
        PhoneNumberAttributedValueField('callback_phones', field_uri='persona:CallbackPhones'),
        PhoneNumberAttributedValueField('car_phones', field_uri='persona:CarPhones'),
        PhoneNumberAttributedValueField('home_faxes', field_uri='persona:HomeFaxes'),
        PhoneNumberAttributedValueField('orgnaization_main_phones', field_uri='persona:OrganizationMainPhones'),
        PhoneNumberAttributedValueField('other_faxes', field_uri='persona:OtherFaxes'),
        PhoneNumberAttributedValueField('other_telephones', field_uri='persona:OtherTelephones'),
        PhoneNumberAttributedValueField('other_phones2', field_uri='persona:OtherPhones2'),
        PhoneNumberAttributedValueField('pagers', field_uri='persona:Pagers'),
        PhoneNumberAttributedValueField('radio_phones', field_uri='persona:RadioPhones'),
        PhoneNumberAttributedValueField('telex_numbers', field_uri='persona:TelexNumbers'),
        PhoneNumberAttributedValueField('tty_tdd_phone_numbers', field_uri='persona:TTYTDDPhoneNumbers'),
        PhoneNumberAttributedValueField('work_faxes', field_uri='persona:WorkFaxes'),
        EmailAddressAttributedValueField('emails1', field_uri='persona:Emails1'),
        EmailAddressAttributedValueField('emails2', field_uri='persona:Emails2'),
        EmailAddressAttributedValueField('emails3', field_uri='persona:Emails3'),
        StringAttributedValueField('business_home_pages', field_uri='persona:BusinessHomePages'),
        StringAttributedValueField('personal_home_pages', field_uri='persona:PersonalHomePages'),
        StringAttributedValueField('office_locations', field_uri='persona:OfficeLocations'),
        StringAttributedValueField('im_addresses', field_uri='persona:ImAddresses'),
        StringAttributedValueField('im_addresses2', field_uri='persona:ImAddresses2'),
        StringAttributedValueField('im_addresses3', field_uri='persona:ImAddresses3'),
        PostalAddressAttributedValueField('business_addresses', field_uri='persona:BusinessAddresses'),
        PostalAddressAttributedValueField('home_addresses', field_uri='persona:HomeAddresses'),
        PostalAddressAttributedValueField('other_addresses', field_uri='persona:OtherAddresses'),
        StringAttributedValueField('titles', field_uri='persona:Titles'),
        StringAttributedValueField('departments', field_uri='persona:Departments'),
        StringAttributedValueField('company_names', field_uri='persona:CompanyNames'),
        StringAttributedValueField('managers', field_uri='persona:Managers'),
        StringAttributedValueField('assistant_names', field_uri='persona:AssistantNames'),
        StringAttributedValueField('professions', field_uri='persona:Professions'),
        StringAttributedValueField('spouse_names', field_uri='persona:SpouseNames'),
        StringAttributedValueField('children', field_uri='persona:Children'),
        StringAttributedValueField('schools', field_uri='persona:Schools'),
        StringAttributedValueField('hobbies', field_uri='persona:Hobbies'),
        StringAttributedValueField('wedding_anniversaries', field_uri='persona:WeddingAnniversaries'),
        StringAttributedValueField('birthdays', field_uri='persona:Birthdays'),
        StringAttributedValueField('locations', field_uri='persona:Locations'),
        # ExtendedPropertyAttributedValueField('extended_properties', field_uri='persona:ExtendedProperties'),
    )
    FIELDS = IdChangeKeyMixIn.FIELDS + LOCAL_FIELDS

    __slots__ = tuple(f.name for f in LOCAL_FIELDS)


class DistributionList(Item):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/distributionlist"""
    ELEMENT_NAME = 'DistributionList'
    LOCAL_FIELDS = Fields(
        CharField('display_name', field_uri='contacts:DisplayName', is_required=True),
        CharField('file_as', field_uri='contacts:FileAs', is_read_only=True),
        ChoiceField('contact_source', field_uri='contacts:ContactSource', choices={
            Choice('Store'), Choice('ActiveDirectory')
        }, is_read_only=True),
        MemberListField('members', field_uri='distributionlist:Members'),
    )
    FIELDS = Item.FIELDS + LOCAL_FIELDS

    __slots__ = tuple(f.name for f in LOCAL_FIELDS)
