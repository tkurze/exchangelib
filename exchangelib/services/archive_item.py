from ..util import create_element, MNS
from ..version import EXCHANGE_2013
from .common import EWSAccountService, create_folder_ids_element, create_item_ids_element


class ArchiveItem(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/archiveitem-operation"""
    SERVICE_NAME = 'ArchiveItem'
    element_container_name = '{%s}Items' % MNS
    supported_from = EXCHANGE_2013

    def call(self, items, to_folder):
        """Move a list of items to a specific folder in the archive mailbox.

        Args:
          items: a list of (id, changekey) tuples or Item objects
          to_folder:

        Returns:
          None

        """
        return self._chunked_get_elements(self.get_payload, items=items, to_folder=to_folder)

    def get_payload(self, items, to_folder):
        archiveitem = create_element('m:%s' % self.SERVICE_NAME)
        folder_id = create_folder_ids_element(tag='m:ArchiveSourceFolderId', folders=[to_folder],
                                              version=self.account.version)
        item_ids = create_item_ids_element(items=items, version=self.account.version)
        archiveitem.append(folder_id)
        archiveitem.append(item_ids)
        return archiveitem
