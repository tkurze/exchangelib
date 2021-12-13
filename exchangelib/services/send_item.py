from .common import EWSAccountService, create_item_ids_element
from ..folders import BaseFolder
from ..properties import FolderId
from ..util import create_element, set_xml_value


class SendItem(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/senditem-operation"""

    SERVICE_NAME = 'SendItem'
    returns_elements = False

    def call(self, items, saved_item_folder):
        if saved_item_folder and not isinstance(saved_item_folder, (BaseFolder, FolderId)):
            raise ValueError(f"'saved_item_folder' {saved_item_folder!r} must be a Folder or FolderId instance")
        return self._chunked_get_elements(self.get_payload, items=items, saved_item_folder=saved_item_folder)

    def get_payload(self, items, saved_item_folder):
        payload = create_element(f'm:{self.SERVICE_NAME}', attrs=dict(SaveItemToFolder=bool(saved_item_folder)))
        payload.append(create_item_ids_element(items=items, version=self.account.version))
        if saved_item_folder:
            payload.append(
                set_xml_value(create_element('m:SavedItemFolderId'), saved_item_folder, version=self.account.version)
            )
        return payload
