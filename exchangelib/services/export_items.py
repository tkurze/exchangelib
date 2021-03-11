from .common import EWSAccountService, create_item_ids_element
from ..errors import ResponseMessageError
from ..util import create_element, MNS


class ExportItems(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/exportitems-operation"""

    ERRORS_TO_CATCH_IN_RESPONSE = ResponseMessageError
    SERVICE_NAME = 'ExportItems'
    element_container_name = '{%s}Data' % MNS

    def call(self, items):
        return self._chunked_get_elements(self.get_payload, items=items)

    def get_payload(self, items):
        exportitems = create_element('m:%s' % self.SERVICE_NAME)
        item_ids = create_item_ids_element(items=items, version=self.account.version)
        exportitems.append(item_ids)
        return exportitems

    # We need to override this since ExportItemsResponseMessage is formatted a
    #  little bit differently. Namely, all we want is the 64bit string in the
    #  Data tag.
    @classmethod
    def _get_elements_in_container(cls, container):
        return [container.text]
