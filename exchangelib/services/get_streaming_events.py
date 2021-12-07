import logging

from .common import EWSAccountService, add_xml_child
from ..errors import EWSError
from ..properties import Notification
from ..util import create_element, get_xml_attr, get_xml_attrs, MNS, DocumentYielder, DummyResponse

log = logging.getLogger(__name__)
xml_log = logging.getLogger('%s.xml' % __name__)


class GetStreamingEvents(EWSAccountService):
    """MSDN:
    https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getstreamingevents-operation
    """

    SERVICE_NAME = 'GetStreamingEvents'
    element_container_name = '{%s}Notifications' % MNS
    prefer_affinity = True

    # Connection status values
    OK = 'OK'
    CLOSED = 'Closed'

    def __init__(self, *args, **kwargs):
        # These values are set each time call() is consumed
        self.connection_status = None
        super().__init__(*args, **kwargs)
        self.streaming = True

    def call(self, subscription_ids, connection_timeout):
        if connection_timeout < 1:
            raise ValueError("'connection_timeout' must be a positive integer")
        return self._elems_to_objs(self._get_elements(payload=self.get_payload(
                subscription_ids=subscription_ids, connection_timeout=connection_timeout,
        )))

    def _elems_to_objs(self, elems):
        for elem in elems:
            if isinstance(elem, Exception):
                yield elem
                continue
            yield Notification.from_xml(elem=elem, account=None)

    @classmethod
    def _get_soap_parts(cls, response, **parse_opts):
        # Pass the response unaltered. We want to use our custom document yielder
        return None, response

    def _get_soap_messages(self, body, **parse_opts):
        # 'body' is actually the raw response passed on by '_get_soap_parts'. We want to continuously read the content,
        # looking for complete XML documents. When we have a full document, we want to parse it as if it was a normal
        # XML response.
        r = body
        for i, doc in enumerate(DocumentYielder(r.iter_content()), start=1):
            xml_log.debug('''Response XML (docs received: %(i)s): %(xml_response)s''', dict(i=i, xml_response=doc))
            response = DummyResponse(url=None, headers=None, request_headers=None, content=doc)
            try:
                _, body = super()._get_soap_parts(response=response, **parse_opts)
            except Exception:
                r.close()  # Release memory
                raise
            # TODO: We're skipping ._update_api_version() here because we don't have access to the 'api_version' used.
            # TODO: We should be doing a lot of error handling for ._get_soap_messages().
            yield from super()._get_soap_messages(body=body, **parse_opts)
            if self.connection_status == self.CLOSED:
                # Don't wait for the TCP connection to timeout
                break

    def _get_element_container(self, message, name=None):
        error_ids_elem = message.find('{%s}ErrorSubscriptionIds' % MNS)
        error_ids = [] if error_ids_elem is None else get_xml_attrs(error_ids_elem, '{%s}SubscriptionId' % MNS)
        self.connection_status = get_xml_attr(message, '{%s}ConnectionStatus' % MNS)  # Either 'OK' or 'Closed'
        log.debug('Connection status is: %s', self.connection_status)
        # Upstream normally expects to find a 'name' tag but our response does not always have it. We still want to
        # call upstream, to have exceptions raised. Return an empty list if there is no 'name' tag and no errors.
        if message.find(name) is None:
            name = None
        try:
            res = super()._get_element_container(message=message, name=name)
        except EWSError as e:
            # When the request contains a combination of good and failing subscription IDs, notifications for the good
            # subscriptions seem to never be returned even though the XML spec allows it. This means there's no point in
            # trying to collect any notifications here and delivering a combination of errors and return values.
            if error_ids:
                e.value += ' (subscription IDs: %s)' % ', '.join(repr(i) for i in error_ids)
            raise e
        return [] if name is None else res

    def get_payload(self, subscription_ids, connection_timeout):
        getstreamingevents = create_element('m:%s' % self.SERVICE_NAME)
        subscriptions_elem = create_element('m:SubscriptionIds')
        for subscription_id in subscription_ids:
            add_xml_child(subscriptions_elem, 't:SubscriptionId', subscription_id)
        if not len(subscriptions_elem):
            raise ValueError('"subscription_ids" must not be empty')

        getstreamingevents.append(subscriptions_elem)
        add_xml_child(getstreamingevents, 'm:ConnectionTimeout', connection_timeout)
        return getstreamingevents
