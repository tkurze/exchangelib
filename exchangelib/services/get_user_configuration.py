from .common import EWSAccountService
from ..properties import UserConfiguration
from ..util import create_element, set_xml_value

ID = 'Id'
DICTIONARY = 'Dictionary'
XML_DATA = 'XmlData'
BINARY_DATA = 'BinaryData'
ALL = 'All'
PROPERTIES_CHOICES = {ID, DICTIONARY, XML_DATA, BINARY_DATA, ALL}


class GetUserConfiguration(EWSAccountService):
    """MSDN:
    https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getuserconfiguration-operation
    """

    SERVICE_NAME = 'GetUserConfiguration'

    def call(self, user_configuration_name, properties):
        if properties not in PROPERTIES_CHOICES:
            raise ValueError("'properties' %r must be one of %s" % (properties, PROPERTIES_CHOICES))
        for elem in self._get_elements(payload=self.get_payload(
                user_configuration_name=user_configuration_name, properties=properties
        )):
            if isinstance(elem, Exception):
                yield elem
                continue
            yield UserConfiguration.from_xml(elem=elem, account=self.account)

    @classmethod
    def _get_elements_in_container(cls, container):
        return container.findall(UserConfiguration.response_tag())

    def get_payload(self, user_configuration_name, properties):
        getuserconfiguration = create_element('m:%s' % self.SERVICE_NAME)
        set_xml_value(getuserconfiguration, user_configuration_name, version=self.account.version)
        user_configuration_properties = create_element('m:UserConfigurationProperties')
        set_xml_value(user_configuration_properties, properties, version=self.account.version)
        getuserconfiguration.append(user_configuration_properties)
        return getuserconfiguration
