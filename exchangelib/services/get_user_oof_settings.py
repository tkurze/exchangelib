from ..util import create_element, set_xml_value, MNS, TNS
from .common import EWSAccountService


class GetUserOofSettings(EWSAccountService):
    """Get automatic reply settings for the specified mailbox.
    MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getuseroofsettings-operation

    """
    SERVICE_NAME = 'GetUserOofSettings'
    element_container_name = '{%s}OofSettings' % TNS

    def call(self, mailbox):
        return self._get_elements(payload=self.get_payload(mailbox=mailbox))

    def get_payload(self, mailbox):
        from ..properties import AvailabilityMailbox
        payload = create_element('m:%sRequest' % self.SERVICE_NAME)
        return set_xml_value(payload, AvailabilityMailbox.from_mailbox(mailbox), version=self.account.version)

    def _get_elements_in_response(self, response):
        from ..settings import OofSettings
        for c in super()._get_elements_in_response(response=response):
            yield OofSettings.from_xml(elem=c, account=self.account)

    @classmethod
    def _get_elements_in_container(cls, container):
        # This service only returns one result, directly in 'container'
        return [container]

    def _get_element_container(self, message, name=None):
        # This service returns the result container outside the response message
        super()._get_element_container(message=message.find(self._response_message_tag()), name=None)
        return message.find(name)

    @staticmethod
    def _response_message_tag():
        return '{%s}ResponseMessage' % MNS
