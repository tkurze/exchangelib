from .common import EWSService
from ..properties import Room
from ..util import create_element, set_xml_value, MNS
from ..version import EXCHANGE_2010


class GetRooms(EWSService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getrooms"""

    SERVICE_NAME = 'GetRooms'
    element_container_name = '{%s}Rooms' % MNS
    supported_from = EXCHANGE_2010

    def call(self, roomlist):
        for elem in self._get_elements(payload=self.get_payload(roomlist=roomlist)):
            if isinstance(elem, Exception):
                yield elem
                continue
            yield Room.from_xml(elem=elem, account=None)

    def get_payload(self, roomlist):
        getrooms = create_element('m:%s' % self.SERVICE_NAME)
        set_xml_value(getrooms, roomlist, version=self.protocol.version)
        return getrooms
