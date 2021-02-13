from ..util import create_element
from .common import EWSAccountService, EWSPooledMixIn, create_item_ids_element


class MarkAsJunk(EWSAccountService, EWSPooledMixIn):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/markasjunk"""
    SERVICE_NAME = 'MarkAsJunk'

    def call(self, items, is_junk, move_item):
        return self._pool_requests(payload_func=self.get_payload, **dict(
            items=items, is_junk=is_junk, move_item=move_item
        ))

    def _get_element_container(self, message, response_message=None, name=None):
        # MarkAsJunk returns MovedItemIds directly beneath MarkAsJunkResponseMessage. Collect the elements
        # and make our own fake container.
        from ..properties import MovedItemId
        res = super()._get_element_container(
            message=message, response_message=response_message, name=name
        )
        if not res:
            return res
        fake_elem = create_element('FakeContainer')
        for elem in message.findall(MovedItemId.response_tag()):
            fake_elem.append(elem)
        return fake_elem

    def get_payload(self, items, is_junk, move_item):
        # Takes a list of items and returns either success or raises an error message
        mark_as_junk = create_element(
            'm:%s' % self.SERVICE_NAME,
            attrs=dict(IsJunk='true' if is_junk else 'false', MoveItem='true' if move_item else 'false')
        )
        item_ids = create_item_ids_element(items=items, version=self.account.version)
        mark_as_junk.append(item_ids)
        return mark_as_junk
