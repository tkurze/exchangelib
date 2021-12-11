from .common import EWSAccountService, to_item_id
from ..properties import PersonaId
from ..util import create_element, set_xml_value, MNS


class GetPersona(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getpersona-operation"""

    SERVICE_NAME = 'GetPersona'

    def call(self, persona):
        return self._elems_to_objs(self._get_elements(payload=self.get_payload(persona=persona)))

    def _elems_to_objs(self, elems):
        from ..items import Persona
        elements = list(elems)
        if len(elements) != 1:
            raise ValueError('Expected exactly one element in response')
        elem = elements[0]
        if isinstance(elem, Exception):
            raise elem
        return Persona.from_xml(elem=elem, account=None)

    def get_payload(self, persona):
        return set_xml_value(
            create_element(f'm:{self.SERVICE_NAME}'),
            to_item_id(persona, PersonaId, version=self.protocol.version),
            version=self.protocol.version
        )

    @classmethod
    def _get_elements_in_container(cls, container):
        from ..items import Persona
        return container.findall(f'{{{MNS}}}{Persona.ELEMENT_NAME}')

    @classmethod
    def _response_tag(cls):
        return f'{{{MNS}}}{cls.SERVICE_NAME}ResponseMessage'
