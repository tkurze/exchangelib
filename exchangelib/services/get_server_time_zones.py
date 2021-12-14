import datetime

from .common import EWSService
from ..errors import NaiveDateTimeNotAllowed
from ..ewsdatetime import EWSDateTime
from ..fields import WEEKDAY_NAMES
from ..util import create_element, set_xml_value, xml_text_to_value, peek, TNS, MNS
from ..version import EXCHANGE_2010


class GetServerTimeZones(EWSService):
    """MSDN:
    https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getservertimezones-operation
    """

    SERVICE_NAME = 'GetServerTimeZones'
    element_container_name = f'{{{MNS}}}TimeZoneDefinitions'
    supported_from = EXCHANGE_2010

    def call(self, timezones=None, return_full_timezone_data=False):
        return self._elems_to_objs(self._get_elements(payload=self.get_payload(
            timezones=timezones,
            return_full_timezone_data=return_full_timezone_data
        )))

    def get_payload(self, timezones, return_full_timezone_data):
        payload = create_element(
            f'm:{self.SERVICE_NAME}',
            attrs=dict(ReturnFullTimeZoneData=return_full_timezone_data),
        )
        if timezones is not None:
            is_empty, timezones = peek(timezones)
            if not is_empty:
                tz_ids = create_element('m:Ids')
                for timezone in timezones:
                    tz_id = set_xml_value(create_element('t:Id'), timezone.ms_id)
                    tz_ids.append(tz_id)
                payload.append(tz_ids)
        return payload

    def _elem_to_obj(self, elem):
        tz_id = elem.get('Id')
        tz_name = elem.get('Name')
        tz_periods = self._get_periods(elem)
        tz_transitions_groups = self._get_transitions_groups(elem)
        tz_transitions = self._get_transitions(elem)
        return tz_id, tz_name, tz_periods, tz_transitions, tz_transitions_groups

    @staticmethod
    def _get_periods(timezone_def):
        tz_periods = {}
        periods = timezone_def.find(f'{{{TNS}}}Periods')
        for period in periods.findall(f'{{{TNS}}}Period'):
            # Convert e.g. "trule:Microsoft/Registry/W. Europe Standard Time/2006-Daylight" to (2006, 'Daylight')
            p_year, p_type = period.get('Id').rsplit('/', 1)[1].split('-')
            tz_periods[(int(p_year), p_type)] = dict(
                name=period.get('Name'),
                bias=xml_text_to_value(period.get('Bias'), datetime.timedelta)
            )
        return tz_periods

    @staticmethod
    def _get_transitions_groups(timezone_def):
        tz_transitions_groups = {}
        transition_groups = timezone_def.find(f'{{{TNS}}}TransitionsGroups')
        if transition_groups is not None:
            for transition_group in transition_groups.findall(f'{{{TNS}}}TransitionsGroup'):
                tg_id = int(transition_group.get('Id'))
                tz_transitions_groups[tg_id] = []
                for transition in transition_group.findall(f'{{{TNS}}}Transition'):
                    # Apply same conversion to To as for period IDs
                    to_year, to_type = transition.find(f'{{{TNS}}}To').text.rsplit('/', 1)[1].split('-')
                    tz_transitions_groups[tg_id].append(dict(
                        to=(int(to_year), to_type),
                    ))
                for transition in transition_group.findall(f'{{{TNS}}}RecurringDayTransition'):
                    # Apply same conversion to To as for period IDs
                    to_year, to_type = transition.find(f'{{{TNS}}}To').text.rsplit('/', 1)[1].split('-')
                    occurrence = xml_text_to_value(transition.find(f'{{{TNS}}}Occurrence').text, int)
                    if occurrence == -1:
                        # See TimeZoneTransition.from_xml()
                        occurrence = 5
                    tz_transitions_groups[tg_id].append(dict(
                        to=(int(to_year), to_type),
                        offset=xml_text_to_value(transition.find(f'{{{TNS}}}TimeOffset').text, datetime.timedelta),
                        iso_month=xml_text_to_value(transition.find(f'{{{TNS}}}Month').text, int),
                        iso_weekday=WEEKDAY_NAMES.index(transition.find(f'{{{TNS}}}DayOfWeek').text) + 1,
                        occurrence=occurrence,
                    ))
        return tz_transitions_groups

    @staticmethod
    def _get_transitions(timezone_def):
        tz_transitions = {}
        transitions = timezone_def.find(f'{{{TNS}}}Transitions')
        if transitions is not None:
            for transition in transitions.findall(f'{{{TNS}}}Transition'):
                to = transition.find(f'{{{TNS}}}To')
                if to.get('Kind') != 'Group':
                    raise ValueError(f"Unexpected 'Kind' XML attr: {to.get('Kind')}")
                tg_id = xml_text_to_value(to.text, int)
                tz_transitions[tg_id] = None
            for transition in transitions.findall(f'{{{TNS}}}AbsoluteDateTransition'):
                to = transition.find(f'{{{TNS}}}To')
                if to.get('Kind') != 'Group':
                    raise ValueError(f"Unexpected 'Kind' XML attr: {to.get('Kind')}")
                tg_id = xml_text_to_value(to.text, int)
                try:
                    t_date = xml_text_to_value(transition.find(f'{{{TNS}}}DateTime').text, EWSDateTime).date()
                except NaiveDateTimeNotAllowed as e:
                    # We encountered a naive datetime. Don't worry. we just need the date
                    t_date = e.local_dt.date()
                tz_transitions[tg_id] = t_date
        return tz_transitions
