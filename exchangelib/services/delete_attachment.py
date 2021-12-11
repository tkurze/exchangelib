from .common import EWSAccountService, create_attachment_ids_element
from ..properties import RootItemId
from ..util import create_element, set_xml_value


class DeleteAttachment(EWSAccountService):
    """MSDN:
    https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteattachment-operation
    """

    SERVICE_NAME = 'DeleteAttachment'

    def call(self, items):
        return self._elems_to_objs(self._chunked_get_elements(self.get_payload, items=items))

    def _elems_to_objs(self, elems):
        for elem in elems:
            if isinstance(elem, Exception):
                yield elem
                continue
            yield RootItemId.from_xml(elem=elem, account=self.account)

    @classmethod
    def _get_elements_in_container(cls, container):
        return container.findall(RootItemId.response_tag())

    def get_payload(self, items):
        return set_xml_value(
            create_element(f'm:{self.SERVICE_NAME}'),
            create_attachment_ids_element(items=items, version=self.account.version),
            version=self.account.version
        )
