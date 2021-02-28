import requests_mock

from exchangelib.errors import ErrorServerBusy, ErrorNonExistentMailbox, TransportError, MalformedResponseError, \
    ErrorInvalidServerVersion, ErrorTooManyObjectsOpened, SOAPError
from exchangelib.folders import FolderCollection
from exchangelib.protocol import FaultTolerance
from exchangelib.services import GetServerTimeZones, GetRoomLists, GetRooms, ResolveNames, FindFolder
from exchangelib.util import create_element
from exchangelib.version import EXCHANGE_2007, EXCHANGE_2010

from .common import EWSTest, mock_protocol, mock_version, mock_account, MockResponse, get_random_string


class ServicesTest(EWSTest):
    def test_invalid_server_version(self):
        # Test that we get a client-side error if we call a service that was only implemented in a later version
        version = mock_version(build=EXCHANGE_2007)
        account = mock_account(version=version, protocol=mock_protocol(version=version, service_endpoint='example.com'))
        with self.assertRaises(NotImplementedError):
            list(GetServerTimeZones(protocol=account.protocol).call())
        with self.assertRaises(NotImplementedError):
            list(GetRoomLists(protocol=account.protocol).call())
        with self.assertRaises(NotImplementedError):
            list(GetRooms(protocol=account.protocol).call('XXX'))

    def test_error_server_busy(self):
        # Test that we can parse an exception response via SOAP body
        version = mock_version(build=EXCHANGE_2010)
        ws = GetRoomLists(mock_protocol(version=version, service_endpoint='example.com'))
        xml = b'''\
<?xml version='1.0' encoding='utf-8'?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <s:Fault>
      <faultcode xmlns:a="http://schemas.microsoft.com/exchange/services/2006/types">a:ErrorServerBusy</faultcode>
      <faultstring xml:lang="en-US">The server cannot service this request right now. Try again later.</faultstring>
      <detail>
        <e:ResponseCode xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">
          ErrorServerBusy
        </e:ResponseCode>
        <e:Message xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">
          The server cannot service this request right now. Try again later.
        </e:Message>
        <t:MessageXml xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
          <t:Value Name="BackOffMilliseconds">297749</t:Value>
        </t:MessageXml>
      </detail>
    </s:Fault>
  </s:Body>
</s:Envelope>'''
        header, body = ws._get_soap_parts(response=MockResponse(xml))
        with self.assertRaises(ErrorServerBusy) as e:
            ws._get_soap_messages(body=body)
        self.assertEqual(e.exception.back_off, 297.749)  # Test that we correctly parse the BackOffMilliseconds value

    @requests_mock.mock(real_http=True)
    def test_error_too_many_objects_opened(self, m):
        # Test that we can parse ErrorTooManyObjectsOpened via ResponseMessage and return
        version = mock_version(build=EXCHANGE_2010)
        protocol = mock_protocol(version=version, service_endpoint='example.com')
        account = mock_account(version=version, protocol=protocol)
        ws = FindFolder(account=account)
        xml = b'''\
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:FindFolderResponse
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
                <m:FindFolderResponseMessage ResponseClass="Error">
                    <m:MessageText>Too many concurrent connections opened.</m:MessageText>
                    <m:ResponseCode>ErrorTooManyObjectsOpened</m:ResponseCode>
                    <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
                </m:FindFolderResponseMessage>
            </m:ResponseMessages>
        </m:FindFolderResponse>
    </s:Body>
</s:Envelope>'''
        header, body = ws._get_soap_parts(response=MockResponse(xml))
        # Just test that we can parse the error
        with self.assertRaises(ErrorTooManyObjectsOpened):
            list(ws._get_elements_in_response(response=ws._get_soap_messages(body=body)))

        # Test that it gets converted to an ErrorServerBusy exception. This happens deep inside EWSService methods
        # so it's easier to only mock the response.
        self.account.root  # Needed to get past the GetFolder request
        m.post(self.account.protocol.service_endpoint, content=xml)
        self.account.protocol.config.retry_policy = FaultTolerance(max_wait=0)
        with self.assertRaises(ErrorServerBusy) as e:
            list(FolderCollection(account=self.account, folders=[self.account.root]).find_folders())
        self.assertEqual(e.exception.back_off, None)  # ErrorTooManyObjectsOpened has no BackOffMilliseconds value

    def test_soap_error(self):
        soap_xml = """\
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <s:Fault>
      <faultcode>{faultcode}</faultcode>
      <faultstring>{faultstring}</faultstring>
      <faultactor>https://CAS01.example.com/EWS/Exchange.asmx</faultactor>
      <detail>
        <ResponseCode xmlns="http://schemas.microsoft.com/exchange/services/2006/errors">{responsecode}</ResponseCode>
        <Message xmlns="http://schemas.microsoft.com/exchange/services/2006/errors">{message}</Message>
      </detail>
    </s:Fault>
  </s:Body>
</s:Envelope>"""
        header, body = ResolveNames._get_soap_parts(response=MockResponse(soap_xml.format(
                faultcode='YYY', faultstring='AAA', responsecode='XXX', message='ZZZ'
            ).encode('utf-8')))
        with self.assertRaises(SOAPError) as e:
            ResolveNames._get_soap_messages(body=body)
        self.assertIn('AAA', e.exception.args[0])
        self.assertIn('YYY', e.exception.args[0])
        self.assertIn('ZZZ', e.exception.args[0])
        header, body = ResolveNames._get_soap_parts(response=MockResponse(soap_xml.format(
                faultcode='ErrorNonExistentMailbox', faultstring='AAA', responsecode='XXX', message='ZZZ'
            ).encode('utf-8')))
        with self.assertRaises(ErrorNonExistentMailbox) as e:
            ResolveNames._get_soap_messages(body=body)
        self.assertIn('AAA', e.exception.args[0])
        header, body = ResolveNames._get_soap_parts(response=MockResponse(soap_xml.format(
                faultcode='XXX', faultstring='AAA', responsecode='ErrorNonExistentMailbox', message='YYY'
            ).encode('utf-8')))
        with self.assertRaises(ErrorNonExistentMailbox) as e:
            ResolveNames._get_soap_messages(body=body)
        self.assertIn('YYY', e.exception.args[0])

        # Test bad XML (no body)
        soap_xml = b"""\
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <t:ServerVersionInfo MajorVersion="8" MinorVersion="0" MajorBuildNumber="685" MinorBuildNumber="8"
                         xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
  </s:Header>
  </s:Body>
</s:Envelope>"""
        with self.assertRaises(MalformedResponseError):
            ResolveNames._get_soap_parts(response=MockResponse(soap_xml))

        # Test bad XML (no fault)
        soap_xml = b"""\
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <t:ServerVersionInfo MajorVersion="8" MinorVersion="0" MajorBuildNumber="685" MinorBuildNumber="8"
                         xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
  </s:Header>
  <s:Body>
    <s:Fault>
    </s:Fault>
  </s:Body>
</s:Envelope>"""
        header, body = ResolveNames._get_soap_parts(response=MockResponse(soap_xml))
        with self.assertRaises(TransportError):
            ResolveNames._get_soap_messages(body=body)

    def test_element_container(self):
        svc = ResolveNames(self.account.protocol)
        soap_xml = b"""\
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <m:ResolveNamesResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <m:ResponseMessages>
        <m:ResolveNamesResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
        </m:ResolveNamesResponseMessage>
      </m:ResponseMessages>
    </m:ResolveNamesResponse>
  </s:Body>
</s:Envelope>"""
        header, body = svc._get_soap_parts(response=MockResponse(soap_xml))
        resp = svc._get_soap_messages(body=body)
        with self.assertRaises(TransportError) as e:
            # Missing ResolutionSet elements
            list(svc._get_elements_in_response(response=resp))
        self.assertIn('ResolutionSet elements in ResponseMessage', e.exception.args[0])

    def test_get_elements(self):
        # Test that we can handle SOAP-level error messages
        # TODO: The request actually raises ErrorInvalidRequest, but we interpret that to mean a wrong API version and
        # end up throwing ErrorInvalidServerVersion. We should make a more direct test.
        svc = ResolveNames(self.account.protocol)
        with self.assertRaises(ErrorInvalidServerVersion):
            list(svc._get_elements(create_element('XXX')))

    @requests_mock.mock()
    def test_invalid_soap_response(self, m):
        m.post(self.account.protocol.service_endpoint, text='XXX')
        with self.assertRaises(SOAPError):
            self.account.inbox.all().count()

    def test_version_renegotiate(self):
        # Test that we can recover from a wrong API version. This is needed in version guessing and when the
        # autodiscover response returns a wrong server version for the account
        old_version = self.account.version.api_version
        self.account.version.api_version = 'Exchange2016'  # Newer EWS versions require a valid value
        try:
            list(self.account.inbox.filter(subject=get_random_string(16)))
            self.assertEqual(old_version, self.account.version.api_version)
        finally:
            self.account.version.api_version = old_version
