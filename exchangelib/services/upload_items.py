from .common import EWSAccountService
from ..properties import ItemId, ParentFolderId
from ..util import create_element, set_xml_value, add_xml_child, MNS


class UploadItems(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/uploaditems-operation

    TODO: This currently has the existing limitation of only being able to upload items that do not yet exist in the
     database. The full spec also allows actions "Update" and "UpdateOrCreate".
    """

    SERVICE_NAME = 'UploadItems'
    element_container_name = '{%s}ItemId' % MNS

    def call(self, items):
        # _pool_requests expects 'items', not 'data'
        return self._chunked_get_elements(self.get_payload, items=items)

    def get_payload(self, items):
        """Upload given items to given account.

        'items' is an iterable of tuples where the first element is a Folder instance representing the ParentFolder
        that the item will be placed in and the second element is a Data string returned from an ExportItems.
        call.

        :param items:
        """
        uploaditems = create_element('m:%s' % self.SERVICE_NAME)
        itemselement = create_element('m:Items')
        uploaditems.append(itemselement)
        for parent_folder, data_str in items:
            item = create_element('t:Item', attrs=dict(CreateAction='CreateNew'))
            parentfolderid = ParentFolderId(parent_folder.id, parent_folder.changekey)
            set_xml_value(item, parentfolderid, version=self.account.version)
            add_xml_child(item, 't:Data', data_str)
            itemselement.append(item)
        return uploaditems

    @classmethod
    def _get_elements_in_container(cls, container):
        return [(container.get(ItemId.ID_ATTR), container.get(ItemId.CHANGEKEY_ATTR))]
