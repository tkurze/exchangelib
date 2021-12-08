import datetime
import os
import pickle
import socket
import tempfile
import warnings
try:
    import zoneinfo
except ImportError:
    from backports import zoneinfo

import psutil
import requests_mock

from exchangelib.credentials import Credentials
from exchangelib.configuration import Configuration
from exchangelib.items import CalendarItem
from exchangelib.errors import SessionPoolMinSizeReached, ErrorNameResolutionNoResults, ErrorAccessDenied, \
    TransportError, SessionPoolMaxSizeReached, TimezoneDefinitionInvalidForYear
from exchangelib.properties import TimeZone, RoomList, FreeBusyView, AlternateId, ID_FORMATS, EWS_ID, \
    SearchableMailbox, FailedMailbox, Mailbox, DLMailbox
from exchangelib.protocol import Protocol, BaseProtocol, NoVerifyHTTPAdapter, FailFast
from exchangelib.services import GetServerTimeZones, GetRoomLists, GetRooms, ResolveNames, GetSearchableMailboxes
from exchangelib.settings import OofSettings
from exchangelib.transport import NOAUTH, NTLM
from exchangelib.version import Build, Version
from exchangelib.winzone import CLDR_TO_MS_TIMEZONE_MAP

from .common import EWSTest, get_random_datetime_range, get_random_string, RANDOM_DATE_MIN, RANDOM_DATE_MAX


class ProtocolTest(EWSTest):

    def test_pickle(self):
        # Test that we can pickle, repr and str Protocols
        o = Protocol(config=Configuration(
            service_endpoint='https://example.com/Foo.asmx',
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NTLM, version=Version(Build(15, 1)), retry_policy=FailFast()
        ))
        pickled_o = pickle.dumps(o)
        unpickled_o = pickle.loads(pickled_o)
        self.assertIsInstance(unpickled_o, type(o))
        self.assertEqual(repr(o), repr(unpickled_o))
        self.assertEqual(str(o), str(unpickled_o))

    @requests_mock.mock()
    def test_session(self, m):
        protocol = Protocol(config=Configuration(
            service_endpoint='https://example.com/Foo.asmx',
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NTLM, version=Version(Build(15, 1)), retry_policy=FailFast()
        ))
        session = protocol.create_session()
        new_session = protocol.renew_session(session)
        self.assertNotEqual(id(session), id(new_session))

    @requests_mock.mock()
    def test_protocol_instance_caching(self, m):
        # Verify that we get the same Protocol instance for the same combination of (endpoint, credentials)
        user, password = get_random_string(8), get_random_string(8)
        config = Configuration(
            service_endpoint='https://example.com/Foo.asmx', credentials=Credentials(user, password),
            auth_type=NTLM, version=Version(Build(15, 1)), retry_policy=FailFast()
        )

        # Test CachingProtocol.__getitem__
        with self.assertRaises(KeyError):
            _ = Protocol[config]
        base_p = Protocol(config=config)
        self.assertEqual(base_p, Protocol[config][0])

        # Make sure we always return the same item when creating a Protocol with the same endpoint and creds
        for _ in range(10):
            p = Protocol(config=config)
            self.assertEqual(base_p, p)
            self.assertEqual(id(base_p), id(p))
            self.assertEqual(hash(base_p), hash(p))
            self.assertEqual(id(base_p._session_pool), id(p._session_pool))

        # Test CachingProtocol.__delitem__
        del Protocol[config]
        with self.assertRaises(KeyError):
            _ = Protocol[config]

        # Make sure we get a fresh instance after we cleared the cache
        p = Protocol(config=config)
        self.assertNotEqual(base_p, p)

        Protocol.clear_cache()

    def test_close(self):
        # Don't use example.com here - it does not resolve or answer on all ISPs
        proc = psutil.Process()
        ip_addresses = {info[4][0] for info in socket.getaddrinfo(
            'httpbin.org', 80, socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_IP
        )}

        def conn_count():
            return len([p for p in proc.connections() if p.raddr[0] in ip_addresses])

        self.assertGreater(len(ip_addresses), 0)
        protocol = Protocol(config=Configuration(
            service_endpoint='http://httpbin.org',
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NOAUTH, version=Version(Build(15, 1)), retry_policy=FailFast(),
            max_connections=3
        ))
        # Merely getting a session should not create conections
        session = protocol.get_session()
        self.assertEqual(conn_count(), 0)
        # Open one URL - we have 1 connection
        session.get('http://httpbin.org')
        self.assertEqual(conn_count(), 1)
        # Open the same URL - we should still have 1 connection
        session.get('http://httpbin.org')
        self.assertEqual(conn_count(), 1)

        # Open some more connections
        s2 = protocol.get_session()
        s2.get('http://httpbin.org')
        s3 = protocol.get_session()
        s3.get('http://httpbin.org')
        self.assertEqual(conn_count(), 3)

        # Releasing the sessions does not close the connections
        protocol.release_session(session)
        protocol.release_session(s2)
        protocol.release_session(s3)
        self.assertEqual(conn_count(), 3)

        # But closing explicitly does
        protocol.close()
        self.assertEqual(conn_count(), 0)

    def test_decrease_poolsize(self):
        # Test increasing and decreasing the pool size
        max_connections = 3
        protocol = Protocol(config=Configuration(
            service_endpoint='https://example.com/Foo.asmx',
            credentials=Credentials(get_random_string(8), get_random_string(8)),
            auth_type=NTLM, version=Version(Build(15, 1)), retry_policy=FailFast(),
            max_connections=max_connections,
        ))
        self.assertEqual(protocol._session_pool.qsize(), 0)
        self.assertEqual(protocol.session_pool_size, 0)
        protocol.increase_poolsize()
        protocol.increase_poolsize()
        protocol.increase_poolsize()
        with self.assertRaises(SessionPoolMaxSizeReached):
            protocol.increase_poolsize()
        self.assertEqual(protocol._session_pool.qsize(), max_connections)
        self.assertEqual(protocol.session_pool_size, max_connections)
        protocol.decrease_poolsize()
        protocol.decrease_poolsize()
        with self.assertRaises(SessionPoolMinSizeReached):
            protocol.decrease_poolsize()
        self.assertEqual(protocol._session_pool.qsize(), 1)

    def test_get_timezones(self):
        ws = GetServerTimeZones(self.account.protocol)
        data = ws.call()
        self.assertAlmostEqual(len(list(data)), 130, delta=30, msg=data)
        # Test shortcut
        self.assertAlmostEqual(len(list(self.account.protocol.get_timezones())), 130, delta=30, msg=data)
        # Test translation to TimeZone objects
        for _, _, periods, transitions, transitionsgroups in self.account.protocol.get_timezones(
                return_full_timezone_data=True):
            try:
                TimeZone.from_server_timezone(
                    periods=periods, transitions=transitions, transitionsgroups=transitionsgroups, for_year=2018,
                )
            except TimezoneDefinitionInvalidForYear:
                pass

    def test_get_free_busy_info(self):
        tz = self.account.default_timezone
        server_timezones = list(self.account.protocol.get_timezones(return_full_timezone_data=True))
        start = datetime.datetime.now(tz=tz)
        end = datetime.datetime.now(tz=tz) + datetime.timedelta(hours=6)
        accounts = [(self.account, 'Organizer', False)]

        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=[(123, 'XXX', 'XXX')], start=0, end=0)
        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=[(self.account, 'XXX', 'XXX')], start=0, end=0)
        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=[(self.account, 'Organizer', 'XXX')], start=0, end=0)
        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=accounts, start=end, end=start)
        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=accounts, start=start, end=end,
                                                     merged_free_busy_interval='XXX')
        with self.assertRaises(ValueError):
            self.account.protocol.get_free_busy_info(accounts=accounts, start=start, end=end, requested_view='XXX')

        for view_info in self.account.protocol.get_free_busy_info(accounts=accounts, start=start, end=end):
            self.assertIsInstance(view_info, FreeBusyView)
            self.assertIsInstance(view_info.working_hours_timezone, TimeZone)
            ms_id = view_info.working_hours_timezone.to_server_timezone(server_timezones, start.year)
            self.assertIn(ms_id, {t[0] for t in CLDR_TO_MS_TIMEZONE_MAP.values()})

        # Test account as simple email
        for view_info in self.account.protocol.get_free_busy_info(
                accounts=[(self.account.primary_smtp_address, 'Organizer', False)], start=start, end=end
        ):
            self.assertIsInstance(view_info, FreeBusyView)

    def test_get_roomlists(self):
        # The test server is not guaranteed to have any room lists which makes this test less useful
        ws = GetRoomLists(self.account.protocol)
        roomlists = ws.call()
        self.assertEqual(list(roomlists), [])
        # Test shortcut
        self.assertEqual(list(self.account.protocol.get_roomlists()), [])

    def test_get_roomlists_parsing(self):
        # Test static XML since server has no roomlists
        xml = b'''\
<?xml version="1.0" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:GetRoomListsResponse ResponseClass="Success"
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseCode>NoError</m:ResponseCode>
            <m:RoomLists>
                <t:Address>
                    <t:Name>Roomlist</t:Name>
                    <t:EmailAddress>roomlist1@example.com</t:EmailAddress>
                    <t:RoutingType>SMTP</t:RoutingType>
                    <t:MailboxType>PublicDL</t:MailboxType>
                </t:Address>
                <t:Address>
                    <t:Name>Roomlist</t:Name>
                    <t:EmailAddress>roomlist2@example.com</t:EmailAddress>
                    <t:RoutingType>SMTP</t:RoutingType>
                    <t:MailboxType>PublicDL</t:MailboxType>
                </t:Address>
            </m:RoomLists>
        </m:GetRoomListsResponse>
    </s:Body>
</s:Envelope>'''
        ws = GetRoomLists(self.account.protocol)
        self.assertSetEqual(
            {rl.email_address for rl in ws.parse(xml)},
            {'roomlist1@example.com', 'roomlist2@example.com'}
        )

    def test_get_rooms(self):
        # The test server is not guaranteed to have any rooms or room lists which makes this test less useful
        roomlist = RoomList(email_address='my.roomlist@example.com')
        ws = GetRooms(self.account.protocol)
        with self.assertRaises(ErrorNameResolutionNoResults):
            list(ws.call(roomlist=roomlist))
        # Test shortcut
        with self.assertRaises(ErrorNameResolutionNoResults):
            list(self.account.protocol.get_rooms('my.roomlist@example.com'))

    def test_get_rooms_parsing(self):
        # Test static XML since server has no rooms
        xml = b'''\
<?xml version="1.0" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:GetRoomsResponse ResponseClass="Success"
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseCode>NoError</m:ResponseCode>
            <m:Rooms>
                <t:Room>
                    <t:Id>
                        <t:Name>room1</t:Name>
                        <t:EmailAddress>room1@example.com</t:EmailAddress>
                        <t:RoutingType>SMTP</t:RoutingType>
                        <t:MailboxType>Mailbox</t:MailboxType>
                    </t:Id>
                </t:Room>
                <t:Room>
                    <t:Id>
                        <t:Name>room2</t:Name>
                        <t:EmailAddress>room2@example.com</t:EmailAddress>
                        <t:RoutingType>SMTP</t:RoutingType>
                        <t:MailboxType>Mailbox</t:MailboxType>
                    </t:Id>
                </t:Room>
            </m:Rooms>
        </m:GetRoomsResponse>
    </s:Body>
</s:Envelope>'''
        ws = GetRooms(self.account.protocol)
        self.assertSetEqual(
            {r.email_address for r in ws.parse(xml)},
            {'room1@example.com', 'room2@example.com'}
        )

    def test_resolvenames(self):
        with self.assertRaises(ValueError):
            self.account.protocol.resolve_names(names=[], search_scope='XXX')
        with self.assertRaises(ValueError):
            self.account.protocol.resolve_names(names=[], shape='XXX')
        self.assertGreaterEqual(
            self.account.protocol.resolve_names(names=['xxx@example.com']),
            []
        )
        self.assertEqual(
            self.account.protocol.resolve_names(names=[self.account.primary_smtp_address]),
            [Mailbox(email_address=self.account.primary_smtp_address)]
        )
        # Test something that's not an email
        self.assertEqual(
            self.account.protocol.resolve_names(names=['foo\\bar']),
            [ErrorNameResolutionNoResults('No results were found.')]
        )
        # Test return_full_contact_data
        mailbox, contact = self.account.protocol.resolve_names(
            names=[self.account.primary_smtp_address],
            return_full_contact_data=True
        )[0]
        self.assertEqual(
            mailbox,
            Mailbox(email_address=self.account.primary_smtp_address)
        )
        self.assertListEqual(
            [e.email.replace('SMTP:', '') for e in contact.email_addresses if e.label == 'EmailAddress1'],
            [self.account.primary_smtp_address]
        )

    def test_resolvenames_parsing(self):
        # Test static XML since server has no roomlists
        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <m:ResolveNamesResponse
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:ResolveNamesResponseMessage ResponseClass="Warning">
          <m:MessageText>Multiple results were found.</m:MessageText>
          <m:ResponseCode>ErrorNameResolutionMultipleResults</m:ResponseCode>
          <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
          <m:ResolutionSet TotalItemsInView="2" IncludesLastItemInRange="true">
            <t:Resolution>
              <t:Mailbox>
                <t:Name>John Doe</t:Name>
                <t:EmailAddress>anne@example.com</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType>
                <t:MailboxType>Mailbox</t:MailboxType>
              </t:Mailbox>
            </t:Resolution>
            <t:Resolution>
              <t:Mailbox>
                <t:Name>John Deer</t:Name>
                <t:EmailAddress>john@example.com</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType>
                <t:MailboxType>Mailbox</t:MailboxType>
              </t:Mailbox>
            </t:Resolution>
          </m:ResolutionSet>
        </m:ResolveNamesResponseMessage>
      </m:ResponseMessages>
    </m:ResolveNamesResponse>
  </s:Body>
</s:Envelope>'''
        ws = ResolveNames(self.account.protocol)
        ws.return_full_contact_data = False
        self.assertSetEqual(
            {m.email_address for m in ws.parse(xml)},
            {'anne@example.com', 'john@example.com'}
        )

    def test_get_searchable_mailboxes(self):
        # Insufficient privileges for the test account, so let's just test the exception
        with self.assertRaises(ErrorAccessDenied):
            self.account.protocol.get_searchable_mailboxes('non_existent_distro@example.com')

        xml = b'''\
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
   <s:Body xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:GetSearchableMailboxesResponse ResponseClass="Success">
         <m:ResponseCode>NoError</m:ResponseCode>
         <m:SearchableMailboxes>
            <t:SearchableMailbox>
               <t:Guid>33a408fe-2574-4e3b-49f5-5e1e000a3035</t:Guid>
               <t:PrimarySmtpAddress>LOLgroup@contoso.com</t:PrimarySmtpAddress>
               <t:IsExternalMailbox>false</t:IsExternalMailbox>
               <t:ExternalEmailAddress/>
               <t:DisplayName>LOLgroup</t:DisplayName>
               <t:IsMembershipGroup>true</t:IsMembershipGroup>
               <t:ReferenceId>/o=First/ou=Exchange(FYLT)/cn=Recipients/cn=81213b958a0b5295b13b3f02b812bf1bc-LOLgroup</t:ReferenceId>
            </t:SearchableMailbox>
            <t:FailedMailbox>
               <t:Mailbox>FAILgroup@contoso.com</t:Mailbox>
               <t:ErrorCode>123</t:ErrorCode>
               <t:ErrorMessage>Catastrophic Failure</t:ErrorMessage>
               <t:IsArchive>true</t:IsArchive>
            </t:FailedMailbox>
         </m:SearchableMailboxes>
      </m:GetSearchableMailboxesResponse>
   </s:Body>
</s:Envelope>'''
        ws = GetSearchableMailboxes(protocol=self.account.protocol)
        self.assertListEqual(list(ws.parse(xml)), [
            SearchableMailbox(
                guid='33a408fe-2574-4e3b-49f5-5e1e000a3035',
                primary_smtp_address='LOLgroup@contoso.com',
                is_external=False,
                external_email=None,
                display_name='LOLgroup',
                is_membership_group=True,
                reference_id='/o=First/ou=Exchange(FYLT)/cn=Recipients/cn=81213b958a0b5295b13b3f02b812bf1bc-LOLgroup',
            ),
            FailedMailbox(
                mailbox='FAILgroup@contoso.com',
                error_code=123,
                error_message='Catastrophic Failure',
                is_archive=True,
            ),
        ])

    def test_expanddl(self):
        with self.assertRaises(ErrorNameResolutionNoResults):
            self.account.protocol.expand_dl('non_existent_distro@example.com')
        with self.assertRaises(ErrorNameResolutionNoResults):
            self.account.protocol.expand_dl(
                DLMailbox(email_address='non_existent_distro@example.com', mailbox_type='PublicDL')
            )

    def test_oof_settings(self):
        # First, ensure a common starting point
        utc = zoneinfo.ZoneInfo('UTC')
        self.account.oof_settings = OofSettings(
            state=OofSettings.DISABLED,
            start=datetime.datetime.combine(RANDOM_DATE_MIN, datetime.time.min, tzinfo=utc),
            end=datetime.datetime.combine(RANDOM_DATE_MAX, datetime.time.max, tzinfo=utc),
        )

        oof = OofSettings(
            state=OofSettings.ENABLED,
            external_audience='None',
            internal_reply="I'm on holidays. See ya guys!",
            external_reply='Dear Sir, your email has now been deleted.',
        )
        self.account.oof_settings = oof
        self.assertEqual(self.account.oof_settings, oof)

        oof = OofSettings(
            state=OofSettings.ENABLED,
            external_audience='Known',
            internal_reply='XXX',
            external_reply='YYY',
        )
        self.account.oof_settings = oof
        self.assertEqual(self.account.oof_settings, oof)

        # Scheduled duration must not be in the past
        tz = self.account.default_timezone
        start, end = get_random_datetime_range(start_date=datetime.datetime.now(tz).date())
        oof = OofSettings(
            state=OofSettings.SCHEDULED,
            external_audience='Known',
            internal_reply="I'm in the pub. See ya guys!",
            external_reply="I'm having a business dinner in town",
            start=start,
            end=end,
        )
        self.account.oof_settings = oof
        self.assertEqual(self.account.oof_settings, oof)

        oof = OofSettings(
            state=OofSettings.DISABLED,
            start=start,
            end=end,
        )
        self.account.oof_settings = oof
        # TODO: For some reason, disabling OOF does not always work. Don't assert because we want a stable test suite
        if self.account.oof_settings != oof:
            self.skipTest('Disabling OOF did not work')

    def test_oof_settings_validation(self):
        utc = zoneinfo.ZoneInfo('UTC')
        with self.assertRaises(ValueError):
            # Needs a start and end
            OofSettings(
                state=OofSettings.SCHEDULED,
            ).clean(version=None)
        with self.assertRaises(ValueError):
            # Start must be before end
            OofSettings(
                state=OofSettings.SCHEDULED,
                start=datetime.datetime(2100, 12, 1, tzinfo=utc),
                end=datetime.datetime(2100, 11, 1, tzinfo=utc),
            ).clean(version=None)
        with self.assertRaises(ValueError):
            # End must be in the future
            OofSettings(
                state=OofSettings.SCHEDULED,
                start=datetime.datetime(2000, 11, 1, tzinfo=utc),
                end=datetime.datetime(2000, 12, 1, tzinfo=utc),
            ).clean(version=None)
        with self.assertRaises(ValueError):
            # Must have an internal and external reply
            OofSettings(
                state=OofSettings.SCHEDULED,
                start=datetime.datetime(2100, 11, 1, tzinfo=utc),
                end=datetime.datetime(2100, 12, 1, tzinfo=utc),
            ).clean(version=None)

    def test_convert_id(self):
        i = 'AAMkADQyYzZmYmUxLTJiYjItNDg2Ny1iMzNjLTIzYWE1NDgxNmZhNABGAAAAAADUebQDarW2Q7G2Ji8hKofPBwAl9iKCsfCfSa9cmjh' \
            '+JCrCAAPJcuhjAAB0l+JSKvzBRYP+FXGewReXAABj6DrMAAA='
        for fmt in ID_FORMATS:
            res = list(self.account.protocol.convert_ids(
                    [AlternateId(id=i, format=EWS_ID, mailbox=self.account.primary_smtp_address)],
                    destination_format=fmt))
            self.assertEqual(len(res), 1)
            self.assertEqual(res[0].format, fmt)

    def test_sessionpool(self):
        # First, empty the calendar
        start = datetime.datetime(2011, 10, 12, 8, tzinfo=self.account.default_timezone)
        end = datetime.datetime(2011, 10, 12, 10, tzinfo=self.account.default_timezone)
        self.account.calendar.filter(start__lt=end, end__gt=start, categories__contains=self.categories).delete()
        items = []
        for i in range(75):
            subject = f'Test Subject {i}'
            item = CalendarItem(
                start=start,
                end=end,
                subject=subject,
                categories=self.categories,
            )
            items.append(item)
        return_ids = self.account.calendar.bulk_create(items=items)
        self.assertEqual(len(return_ids), len(items))
        ids = self.account.calendar.filter(start__lt=end, end__gt=start, categories__contains=self.categories) \
            .values_list('id', 'changekey')
        self.assertEqual(ids.count(), len(items))

    def test_disable_ssl_verification(self):
        # Test that we can make requests when SSL verification is turned off. I don't know how to mock TLS responses
        if not self.verify_ssl:
            # We can only run this test if we haven't already disabled TLS
            self.skipTest('TLS verification already disabled')

        default_adapter_cls = BaseProtocol.HTTP_ADAPTER_CLS

        # Just test that we can query
        self.account.root.all().exists()

        # Smash TLS verification using an untrusted certificate
        with tempfile.NamedTemporaryFile() as f:
            f.write(b'''\
 -----BEGIN CERTIFICATE-----
MIIENzCCAx+gAwIBAgIJAOYfYfw7NCOcMA0GCSqGSIb3DQEBBQUAMIGxMQswCQYD
VQQGEwJVUzERMA8GA1UECAwITWFyeWxhbmQxFDASBgNVBAcMC0ZvcmVzdCBIaWxs
MScwJQYDVQQKDB5UaGUgQXBhY2hlIFNvZnR3YXJlIEZvdW5kYXRpb24xFjAUBgNV
BAsMDUFwYWNoZSBUaHJpZnQxEjAQBgNVBAMMCWxvY2FsaG9zdDEkMCIGCSqGSIb3
DQEJARYVZGV2QHRocmlmdC5hcGFjaGUub3JnMB4XDTE0MDQwNzE4NTgwMFoXDTIy
MDYyNDE4NTgwMFowgbExCzAJBgNVBAYTAlVTMREwDwYDVQQIDAhNYXJ5bGFuZDEU
MBIGA1UEBwwLRm9yZXN0IEhpbGwxJzAlBgNVBAoMHlRoZSBBcGFjaGUgU29mdHdh
cmUgRm91bmRhdGlvbjEWMBQGA1UECwwNQXBhY2hlIFRocmlmdDESMBAGA1UEAwwJ
bG9jYWxob3N0MSQwIgYJKoZIhvcNAQkBFhVkZXZAdGhyaWZ0LmFwYWNoZS5vcmcw
ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCqE9TE9wEXp5LRtLQVDSGQ
GV78+7ZtP/I/ZaJ6Q6ZGlfxDFvZjFF73seNhAvlKlYm/jflIHYLnNOCySN8I2Xw6
L9MbC+jvwkEKfQo4eDoxZnOZjNF5J1/lZtBeOowMkhhzBMH1Rds351/HjKNg6ZKg
2Cldd0j7HbDtEixOLgLbPRpBcaYrLrNMasf3Hal+x8/b8ue28x93HSQBGmZmMIUw
AinEu/fNP4lLGl/0kZb76TnyRpYSPYojtS6CnkH+QLYnsRREXJYwD1Xku62LipkX
wCkRTnZ5nUsDMX6FPKgjQFQCWDXG/N096+PRUQAChhrXsJ+gF3NqWtDmtrhVQF4n
AgMBAAGjUDBOMB0GA1UdDgQWBBQo8v0wzQPx3EEexJPGlxPK1PpgKjAfBgNVHSME
GDAWgBQo8v0wzQPx3EEexJPGlxPK1PpgKjAMBgNVHRMEBTADAQH/MA0GCSqGSIb3
DQEBBQUAA4IBAQBGFRiJslcX0aJkwZpzTwSUdgcfKbpvNEbCNtVohfQVTI4a/oN5
U+yqDZJg3vOaOuiAZqyHcIlZ8qyesCgRN314Tl4/JQ++CW8mKj1meTgo5YFxcZYm
T9vsI3C+Nzn84DINgI9mx6yktIt3QOKZRDpzyPkUzxsyJ8J427DaimDrjTR+fTwD
1Dh09xeeMnSa5zeV1HEDyJTqCXutLetwQ/IyfmMBhIx+nvB5f67pz/m+Dv6V0r3I
p4HCcdnDUDGJbfqtoqsAATQQWO+WWuswB6mOhDbvPTxhRpZq6AkgWqv4S+u3M2GO
r5p9FrBgavAw5bKO54C0oQKpN/5fta5l6Ws0
-----END CERTIFICATE-----''')
            try:
                os.environ['REQUESTS_CA_BUNDLE'] = f.name
                # Setting the credentials is just an easy way of resetting the session pool. This will let requests
                # pick up the new environment variable. Now the request should fail
                self.account.protocol.credentials = self.account.protocol.credentials
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    # Ignore ResourceWarning for unclosed socket. It does get closed.
                    with self.assertRaises(TransportError) as e:
                        self.account.root.all().exists()
                    self.assertIn('SSLError', e.exception.args[0])

                # Disable insecure TLS warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    # Make sure we can handle TLS validation errors when using the custom adapter
                    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
                    self.account.protocol.credentials = self.account.protocol.credentials
                    self.account.root.all().exists()

                    # Test that the custom adapter also works when validation is OK again
                    del os.environ['REQUESTS_CA_BUNDLE']
                    self.account.protocol.credentials = self.account.protocol.credentials
                    self.account.root.all().exists()
            finally:
                # Reset environment and connections
                os.environ.pop('REQUESTS_CA_BUNDLE', None)  # May already have been deleted
                BaseProtocol.HTTP_ADAPTER_CLS = default_adapter_cls
                self.account.protocol.credentials = self.account.protocol.credentials
