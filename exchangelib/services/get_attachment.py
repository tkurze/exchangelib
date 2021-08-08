from .common import EWSAccountService, create_attachment_ids_element
from ..util import create_element, add_xml_child, DummyResponse, StreamingBase64Parser, StreamingContentHandler, \
    ElementNotFound, MNS


class GetAttachment(EWSAccountService):
    """MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getattachment-operation"""

    SERVICE_NAME = 'GetAttachment'
    element_container_name = '{%s}Attachments' % MNS
    streaming = True

    def call(self, items, include_mime_content):
        return self._elems_to_objs(self._chunked_get_elements(
                self.get_payload, items=items, include_mime_content=include_mime_content
        ))

    def _elems_to_objs(self, elems):
        from ..attachments import FileAttachment, ItemAttachment
        cls_map = {cls.response_tag(): cls for cls in (FileAttachment, ItemAttachment)}
        for elem in elems:
            if isinstance(elem, Exception):
                yield elem
                continue
            yield cls_map[elem.tag].from_xml(elem=elem, account=self.account)

    def get_payload(self, items, include_mime_content):
        payload = create_element('m:%s' % self.SERVICE_NAME)
        # TODO: Support additional properties of AttachmentShape. See
        #  https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachmentshape
        if include_mime_content:
            attachment_shape = create_element('m:AttachmentShape')
            add_xml_child(attachment_shape, 't:IncludeMimeContent', 'true')
            payload.append(attachment_shape)
        attachment_ids = create_attachment_ids_element(items=items, version=self.account.version)
        payload.append(attachment_ids)
        return payload

    def _update_api_version(self, api_version, header, **parse_opts):
        if not parse_opts.get('stream_file_content', False):
            super()._update_api_version(api_version, header, **parse_opts)
        # TODO: We're skipping this part in streaming mode because StreamingBase64Parser cannot parse the SOAP header

    @classmethod
    def _get_soap_parts(cls, response, **parse_opts):
        if not parse_opts.get('stream_file_content', False):
            return super()._get_soap_parts(response, **parse_opts)

        # Pass the response unaltered. We want to use our custom streaming parser
        return None, response

    def _get_soap_messages(self, body, **parse_opts):
        if not parse_opts.get('stream_file_content', False):
            return super()._get_soap_messages(body, **parse_opts)

        from ..attachments import FileAttachment
        # 'body' is actually the raw response passed on by '_get_soap_parts'
        r = body
        parser = StreamingBase64Parser()
        field = FileAttachment.get_field_by_fieldname('_content')
        handler = StreamingContentHandler(parser=parser, ns=field.namespace, element_name=field.field_uri)
        parser.setContentHandler(handler)
        return parser.parse(r)

    def stream_file_content(self, attachment_id):
        # The streaming XML parser can only stream content of one attachment
        payload = self.get_payload(items=[attachment_id], include_mime_content=False)
        try:
            yield from self._get_response_xml(payload=payload, stream_file_content=True)
        except ElementNotFound as enf:
            # When the returned XML does not contain a Content element, ElementNotFound is thrown by parser.parse().
            # Let the non-streaming SOAP parser parse the response and hook into the normal exception handling.
            # Wrap in DummyResponse because _get_soap_parts() expects an iter_content() method.
            response = DummyResponse(url=None, headers=None, request_headers=None, content=enf.data)
            _, body = super()._get_soap_parts(response=response)
            res = super()._get_soap_messages(body=body)
            for e in self._get_elements_in_response(response=res):
                if isinstance(e, Exception):
                    raise e
            # The returned content did not contain any EWS exceptions. Give up and re-raise the original exception.
            raise enf
        finally:
            self.stop_streaming()
