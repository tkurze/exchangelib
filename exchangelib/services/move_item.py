from .common import EWSAccountService, create_item_ids_element
from ..util import create_element, set_xml_value, MNS


class MoveItem(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/moveitem-operation"""

    SERVICE_NAME = 'MoveItem'
    element_container_name = '{%s}Items' % MNS

    def call(self, items, to_folder):
        from ..folders import BaseFolder, FolderId
        from ..items import Item
        if not isinstance(to_folder, (BaseFolder, FolderId)):
            raise ValueError("'to_folder' %r must be a Folder or FolderId instance" % to_folder)
        for elem in self._chunked_get_elements(self.get_payload, items=items, to_folder=to_folder):
            if isinstance(elem, (Exception, type(None))):
                yield elem
                continue
            yield Item.id_from_xml(elem)

    def get_payload(self, items, to_folder):
        # Takes a list of items and returns their new item IDs
        moveitem = create_element('m:%s' % self.SERVICE_NAME)
        tofolderid = create_element('m:ToFolderId')
        set_xml_value(tofolderid, to_folder, version=self.account.version)
        moveitem.append(tofolderid)
        item_ids = create_item_ids_element(items=items, version=self.account.version)
        moveitem.append(item_ids)
        return moveitem
