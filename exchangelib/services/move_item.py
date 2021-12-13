from .common import EWSAccountService, create_item_ids_element
from ..folders import BaseFolder
from ..items import Item
from ..properties import FolderId
from ..util import create_element, set_xml_value, MNS


class MoveItem(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/moveitem-operation"""

    SERVICE_NAME = 'MoveItem'
    element_container_name = f'{{{MNS}}}Items'

    def call(self, items, to_folder):
        if not isinstance(to_folder, (BaseFolder, FolderId)):
            raise ValueError(f"'to_folder' {to_folder!r} must be a Folder or FolderId instance")
        return self._elems_to_objs(self._chunked_get_elements(self.get_payload, items=items, to_folder=to_folder))

    def _elems_to_objs(self, elems):
        for elem in elems:
            if isinstance(elem, (Exception, type(None))):
                yield elem
                continue
            yield Item.id_from_xml(elem)

    def get_payload(self, items, to_folder):
        # Takes a list of items and returns their new item IDs
        payload = create_element(f'm:{self.SERVICE_NAME}')
        payload.append(set_xml_value(create_element('m:ToFolderId'), to_folder, version=self.account.version))
        payload.append(create_item_ids_element(items=items, version=self.account.version))
        return payload
