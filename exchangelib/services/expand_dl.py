from .common import EWSService
from ..errors import ErrorNameResolutionNoResults, ErrorNameResolutionMultipleResults
from ..properties import Mailbox
from ..util import create_element, set_xml_value, MNS


class ExpandDL(EWSService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/expanddl-operation"""

    SERVICE_NAME = 'ExpandDL'
    element_container_name = '{%s}DLExpansion' % MNS
    ERRORS_TO_CATCH_IN_RESPONSE = ErrorNameResolutionNoResults
    WARNINGS_TO_IGNORE_IN_RESPONSE = ErrorNameResolutionMultipleResults

    def call(self, distribution_list):
        for elem in self._get_elements(payload=self.get_payload(distribution_list=distribution_list)):
            if isinstance(elem, Exception):
                raise elem
            yield Mailbox.from_xml(elem, account=None)

    def get_payload(self, distribution_list):
        payload = create_element('m:%s' % self.SERVICE_NAME)
        set_xml_value(payload, distribution_list, version=self.protocol.version)
        return payload
