from collections import namedtuple
from unittest.mock import Mock

import requests_mock

from exchangelib.account import DELEGATE, Identity
from exchangelib.credentials import OAuth2Credentials
from exchangelib.errors import (
    ErrorExceededConnectionCount,
    ErrorInternalServerError,
    ErrorInvalidServerVersion,
    ErrorInvalidValueForProperty,
    ErrorNonExistentMailbox,
    ErrorSchemaValidation,
    ErrorServerBusy,
    ErrorTooManyObjectsOpened,
    MalformedResponseError,
    RateLimitError,
    SOAPError,
    TransportError,
)
from exchangelib.folders import FolderCollection
from exchangelib.protocol import FailFast, FaultTolerance
from exchangelib.services import DeleteItem, FindFolder, GetRoomLists, GetRooms, GetServerTimeZones, ResolveNames
from exchangelib.services.common import EWSAccountService, EWSService
from exchangelib.util import PrettyXmlHandler, create_element
from exchangelib.version import EXCHANGE_2007, EXCHANGE_2010

from .common import EWSTest, get_random_string, mock_account, mock_protocol, mock_version


class ServicesTest(EWSTest):
    def test_invalid_server_version(self):
        # Test that we get a client-side error if we call a service that was only implemented in a later version
        version = mock_version(build=EXCHANGE_2007)
        account = mock_account(version=version, protocol=mock_protocol(version=version, service_endpoint="example.com"))
        with self.assertRaises(NotImplementedError):
            list(GetServerTimeZones(protocol=account.protocol).call())
        with self.assertRaises(NotImplementedError):
            list(GetRoomLists(protocol=account.protocol).call())
        with self.assertRaises(NotImplementedError):
            list(GetRooms(protocol=account.protocol).call("XXX"))

    def test_inner_error_parsing(self):
        # Test that we can parse an exception response via SOAP body
        xml = b"""\
<?xml version='1.0' encoding='utf-8'?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <m:DeleteItemResponse
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:DeleteItemResponseMessage ResponseClass="Error">
          <m:MessageText>An internal server error occurred. The operation failed.</m:MessageText>
          <m:ResponseCode>ErrorInternalServerError</m:ResponseCode>
          <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
          <m:MessageXml>
            <t:Value Name="InnerErrorMessageText">Cannot delete message because the folder is out of quota.</t:Value>
            <t:Value Name="InnerErrorResponseCode">ErrorQuotaExceededOnDelete</t:Value>
            <t:Value Name="InnerErrorDescriptiveLinkKey">0</t:Value>
          </m:MessageXml>
        </m:DeleteItemResponseMessage>
      </m:ResponseMessages>
    </m:DeleteItemResponse>
  </s:Body>
</s:Envelope>"""
        ws = DeleteItem(account=self.account)
        with self.assertRaises(ErrorInternalServerError) as e:
            list(ws.parse(xml))
        self.assertEqual(
            e.exception.args[0],
            "An internal server error occurred. The operation failed. (inner error: "
            "ErrorQuotaExceededOnDelete('Cannot delete message because the folder is out of quota.'))",
        )

    def test_invalid_value_extras(self):
        # Test that we can parse an exception response via SOAP body
        xml = b"""\
<?xml version='1.0' encoding='utf-8'?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <m:DeleteItemResponse
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:DeleteItemResponseMessage ResponseClass="Error">
          <m:MessageText>The specified value is invalid for property.</m:MessageText>
          <m:ResponseCode>ErrorInvalidValueForProperty</m:ResponseCode>
          <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
          <m:MessageXml>
            <t:Value Name="Foo">XXX</t:Value>
            <t:Value Name="Bar">YYY</t:Value>
          </m:MessageXml>
        </m:DeleteItemResponseMessage>
      </m:ResponseMessages>
    </m:DeleteItemResponse>
  </s:Body>
</s:Envelope>"""
        ws = DeleteItem(account=self.account)
        with self.assertRaises(ErrorInvalidValueForProperty) as e:
            list(ws.parse(xml))
        self.assertEqual(e.exception.args[0], "The specified value is invalid for property. (Foo: XXX, Bar: YYY)")

    def test_error_server_busy(self):
        # Test that we can parse an exception response via SOAP body
        xml = b"""\
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
</s:Envelope>"""
        version = mock_version(build=EXCHANGE_2010)
        ws = GetRoomLists(mock_protocol(version=version, service_endpoint="example.com"))
        with self.assertRaises(ErrorServerBusy) as e:
            ws.parse(xml)
        self.assertEqual(e.exception.back_off, 297.749)  # Test that we correctly parse the BackOffMilliseconds value

    def test_error_schema_validation(self):
        # Test that we can parse extra info with ErrorSchemaValidation
        xml = b"""\
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <s:Fault>
            <faultcode xmlns:a="http://schemas.microsoft.com/exchange/services/2006/types">a:ErrorSchemaValidation
            </faultcode>
            <faultstring>XXX</faultstring>
            <detail>
                <e:ResponseCode xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">
                ErrorSchemaValidation</e:ResponseCode>
                <e:Message xmlns:e="http://schemas.microsoft.com/exchange/services/2006/errors">YYY</e:Message>
                <t:MessageXml xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                    <t:LineNumber>123</t:LineNumber>
                    <t:LinePosition>456</t:LinePosition>
                    <t:Violation>ZZZ</t:Violation>
                </t:MessageXml>
            </detail>
        </s:Fault>
    </s:Body>
</s:Envelope>"""
        version = mock_version(build=EXCHANGE_2010)
        ws = GetRoomLists(mock_protocol(version=version, service_endpoint="example.com"))
        with self.assertRaises(ErrorSchemaValidation) as e:
            ws.parse(xml)
        self.assertEqual(e.exception.args[0], "YYY ZZZ (line: 123 position: 456)")

    @requests_mock.mock(real_http=True)
    def test_error_too_many_objects_opened(self, m):
        # Test that we can parse ErrorTooManyObjectsOpened via ResponseMessage and return
        version = mock_version(build=EXCHANGE_2010)
        protocol = mock_protocol(version=version, service_endpoint="example.com")
        account = mock_account(version=version, protocol=protocol)
        ws = FindFolder(account=account)
        xml = b"""\
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
</s:Envelope>"""
        # Just test that we can parse the error
        with self.assertRaises(ErrorTooManyObjectsOpened):
            list(ws.parse(xml))

        # Test that it eventually gets converted to an RateLimitError exception. This happens deep inside EWSService
        # methods, so it's easier to only mock the response.
        self.account.root  # Needed to get past the GetFolder request
        m.post(self.account.protocol.service_endpoint, content=xml)
        orig_policy = self.account.protocol.config.retry_policy
        try:
            self.account.protocol.config.retry_policy = FaultTolerance(max_wait=1)  # Set max_wait < RETRY_WAIT
            with self.assertRaises(RateLimitError) as e:
                list(FolderCollection(account=self.account, folders=[self.account.root]).find_folders())
            self.assertEqual(e.exception.wait, self.account.protocol.RETRY_WAIT)
        finally:
            self.account.protocol.config.retry_policy = orig_policy

    def test_soap_error(self):
        xml_template = """\
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
        version = mock_version(build=EXCHANGE_2010)
        protocol = mock_protocol(version=version, service_endpoint="example.com")
        ws = GetRoomLists(protocol=protocol)
        xml = xml_template.format(faultcode="YYY", faultstring="AAA", responsecode="XXX", message="ZZZ").encode("utf-8")
        with self.assertRaises(SOAPError) as e:
            ws.parse(xml)
        self.assertIn("AAA", e.exception.args[0])
        self.assertIn("YYY", e.exception.args[0])
        self.assertIn("ZZZ", e.exception.args[0])
        xml = xml_template.format(
            faultcode="ErrorNonExistentMailbox", faultstring="AAA", responsecode="XXX", message="ZZZ"
        ).encode("utf-8")
        with self.assertRaises(ErrorNonExistentMailbox) as e:
            ws.parse(xml)
        self.assertIn("AAA", e.exception.args[0])
        xml = xml_template.format(
            faultcode="XXX", faultstring="AAA", responsecode="ErrorNonExistentMailbox", message="YYY"
        ).encode("utf-8")
        with self.assertRaises(ErrorNonExistentMailbox) as e:
            ws.parse(xml)
        self.assertIn("YYY", e.exception.args[0])

        # Test bad XML (no body)
        xml = b"""\
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <t:ServerVersionInfo MajorVersion="8" MinorVersion="0" MajorBuildNumber="685" MinorBuildNumber="8"
                         xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
  </s:Header>
  </s:Body>
</s:Envelope>"""
        with self.assertRaises(MalformedResponseError):
            ws.parse(xml)

        # Test bad XML (no fault)
        xml = b"""\
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
        with self.assertRaises(SOAPError) as e:
            ws.parse(xml)
        self.assertEqual(e.exception.args[0], "SOAP error code: None string: None actor: None detail: None")

    def test_element_container(self):
        ws = ResolveNames(self.account.protocol)
        xml = b"""\
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
        with self.assertRaises(TransportError) as e:
            # Missing ResolutionSet elements
            list(ws.parse(xml))
        self.assertIn("ResolutionSet elements in ResponseMessage", e.exception.args[0])

    def test_get_elements(self):
        # Test that we can handle SOAP-level error messages
        # TODO: The request actually raises ErrorInvalidRequest, but we interpret that to mean a wrong API version and
        #  end up throwing ErrorInvalidServerVersion. We should make a more direct test.
        svc = ResolveNames(self.account.protocol)
        with self.assertRaises(ErrorInvalidServerVersion):
            list(svc._get_elements(create_element("XXX")))

    def test_handle_backoff(self):
        # Test that we can handle backoff messages
        svc = ResolveNames(self.account.protocol)
        tmp = svc._response_generator
        orig_policy = self.account.protocol.config.retry_policy
        try:
            # We need to fail fast so we don't end up in an infinite loop
            self.account.protocol.config.retry_policy = FailFast()
            svc._response_generator = Mock(side_effect=ErrorServerBusy("XXX", back_off=1))
            with self.assertRaises(ErrorServerBusy) as e:
                list(svc._get_elements(create_element("XXX")))
            self.assertEqual(e.exception.args[0], "XXX")
        finally:
            svc._response_generator = tmp
            self.account.protocol.config.retry_policy = orig_policy

    def test_exceeded_connection_count(self):
        # Test server repeatedly returning ErrorExceededConnectionCount
        svc = ResolveNames(self.account.protocol)
        tmp = svc._get_soap_messages
        try:
            # We need to fail fast so we don't end up in an infinite loop
            svc._get_soap_messages = Mock(side_effect=ErrorExceededConnectionCount("XXX"))
            with self.assertRaises(ErrorExceededConnectionCount) as e:
                list(svc.call(unresolved_entries=["XXX"]))
            self.assertEqual(e.exception.args[0], "XXX")
        finally:
            svc._get_soap_messages = tmp

    @requests_mock.mock()
    def test_invalid_soap_response(self, m):
        m.post(self.account.protocol.service_endpoint, text="XXX")
        with self.assertRaises(SOAPError):
            self.account.inbox.all().count()

    def test_version_renegotiate(self):
        # Test that we can recover from a wrong API version. This is needed in version guessing and when the
        # autodiscover response returns a wrong server version for the account
        old_version = self.account.version.api_version
        try:
            self.account.version.api_version = "Exchange2016"  # Newer EWS versions require a valid value
            list(self.account.inbox.filter(subject=get_random_string(16)))
            self.assertEqual(old_version, self.account.version.api_version)
        finally:
            self.account.version.api_version = old_version

    def test_wrap(self):
        # Test payload wrapper with both delegation, impersonation and timezones
        svc = EWSService(protocol=None)
        wrapped = svc.wrap(content=create_element("AAA"), api_version="BBB")
        self.assertEqual(
            PrettyXmlHandler().prettify_xml(wrapped),
            """\
<?xml version='1.0' encoding='utf-8'?>
<s:Envelope
    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Header>
    <t:RequestServerVersion Version="BBB"/>
  </s:Header>
  <s:Body>
    <AAA/>
  </s:Body>
</s:Envelope>
""".encode(),
        )

        MockTZ = namedtuple("EWSTimeZone", ["ms_id"])
        MockProtocol = namedtuple("Protocol", ["credentials"])
        MockAccount = namedtuple("Account", ["access_type", "default_timezone", "protocol"])
        for attr, tag in (
            ("primary_smtp_address", "PrimarySmtpAddress"),
            ("upn", "PrincipalName"),
            ("sid", "SID"),
            ("smtp_address", "SmtpAddress"),
        ):
            val = f"{attr}@example.com"
            protocol = MockProtocol(
                credentials=OAuth2Credentials(identity=Identity(**{attr: val}), client_id=None, client_secret=None)
            )
            account = MockAccount(access_type=DELEGATE, default_timezone=MockTZ("XXX"), protocol=protocol)
            svc = EWSAccountService(account=account)
            wrapped = svc.wrap(
                content=create_element("AAA"),
                api_version="BBB",
            )
            self.assertEqual(
                PrettyXmlHandler().prettify_xml(wrapped),
                f"""\
<?xml version='1.0' encoding='utf-8'?>
<s:Envelope
    xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Header>
    <t:RequestServerVersion Version="BBB"/>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:{tag}>{val}</t:{tag}>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
    <t:TimeZoneContext>
      <t:TimeZoneDefinition Id="XXX"/>
    </t:TimeZoneContext>
  </s:Header>
  <s:Body>
    <AAA/>
  </s:Body>
</s:Envelope>
""".encode(),
            )
