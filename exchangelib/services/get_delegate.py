from ..util import create_element, set_xml_value, MNS
from ..version import EXCHANGE_2007_SP1
from .common import EWSAccountService


class GetDelegate(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getdelegate-operation"""
    SERVICE_NAME = 'GetDelegate'
    supported_from = EXCHANGE_2007_SP1

    def call(self, user_ids, include_permissions):
        from ..properties import DLMailbox, DelegateUser  # The service expects a Mailbox element in the MNS namespace

        if user_ids:
            # Pool requests to avoid arbitrarily large requests when user_ids is huge
            res = self._chunked_get_elements(
                self.get_payload,
                items=user_ids,
                mailbox=DLMailbox(email_address=self.account.primary_smtp_address),
                include_permissions=include_permissions,
            )
        else:
            # Pooling expects an iterable of items but we have None. Just call _get_elements directly.
            res = self._get_elements(payload=self.get_payload(
                user_ids=user_ids,
                mailbox=DLMailbox(email_address=self.account.primary_smtp_address),
                include_permissions=include_permissions,
            ))

        for elem in res:
            if isinstance(elem, Exception):
                raise elem
            yield DelegateUser.from_xml(elem=elem, account=self.account)

    def get_payload(self, user_ids, mailbox, include_permissions):
        payload = create_element(
            'm:%s' % self.SERVICE_NAME,
            attrs=dict(IncludePermissions='true' if include_permissions else 'false'),
        )
        set_xml_value(payload, mailbox, version=self.protocol.version)
        if user_ids:
            set_xml_value(payload, user_ids, version=self.protocol.version)
        return payload

    @staticmethod
    def _get_elements_in_container(container):
        from ..properties import DelegateUser
        return container.findall(DelegateUser.response_tag())

    @classmethod
    def _response_message_tag(cls):
        return '{%s}DelegateUserResponseMessageType' % MNS
