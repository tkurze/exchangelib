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

    def _get_elements_in_response(self, response):
        for msg in response:
            container_or_exc = self._get_element_container(message=msg, name=self.element_container_name)
            if isinstance(container_or_exc, (bool, Exception)):
                yield container_or_exc
            else:
                if len(container_or_exc):
                    raise ValueError('Unexpected container length: %s' % container_or_exc)
                yield True

    def get_payload(self, items, to_folder):
        archiveitem = create_element('m:%s' % self.SERVICE_NAME)
        folder_id = create_folder_ids_element(tag='m:ArchiveSourceFolderId', folders=[to_folder],
                                              version=self.account.version)
        item_ids = create_item_ids_element(items=items, version=self.account.version)
        archiveitem.append(folder_id)
        archiveitem.append(item_ids)
        return archiveitem
