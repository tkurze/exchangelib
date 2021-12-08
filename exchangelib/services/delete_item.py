from .common import EWSAccountService, create_item_ids_element
from ..util import create_element
from ..version import EXCHANGE_2013_SP1


class DeleteItem(EWSAccountService):
    """Take a folder and a list of (id, changekey) tuples. Return result of deletion as a list of tuples
    (success[True|False], errormessage), in the same order as the input list.

    MSDN: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/deleteitem-operation
    """

    SERVICE_NAME = 'DeleteItem'
    returns_elements = False

    def call(self, items, delete_type, send_meeting_cancellations, affected_task_occurrences, suppress_read_receipts):
        from ..items import DELETE_TYPE_CHOICES, SEND_MEETING_CANCELLATIONS_CHOICES, AFFECTED_TASK_OCCURRENCES_CHOICES
        if delete_type not in DELETE_TYPE_CHOICES:
            raise ValueError(f"'delete_type' {delete_type} must be one of {DELETE_TYPE_CHOICES}")
        if send_meeting_cancellations not in SEND_MEETING_CANCELLATIONS_CHOICES:
            raise ValueError(f"'send_meeting_cancellations' {send_meeting_cancellations} must be one of "
                             f"{SEND_MEETING_CANCELLATIONS_CHOICES}")
        if affected_task_occurrences not in AFFECTED_TASK_OCCURRENCES_CHOICES:
            raise ValueError(f"'affected_task_occurrences' {affected_task_occurrences} must be one of "
                             f"{AFFECTED_TASK_OCCURRENCES_CHOICES}")
        if suppress_read_receipts not in (True, False):
            raise ValueError(f"'suppress_read_receipts' {suppress_read_receipts} must be True or False")
        return self._chunked_get_elements(
            self.get_payload,
            items=items,
            delete_type=delete_type,
            send_meeting_cancellations=send_meeting_cancellations,
            affected_task_occurrences=affected_task_occurrences,
            suppress_read_receipts=suppress_read_receipts,
        )

    def get_payload(self, items, delete_type, send_meeting_cancellations, affected_task_occurrences,
                    suppress_read_receipts):
        # Takes a list of (id, changekey) tuples or Item objects and returns the XML for a DeleteItem request.
        if self.account.version.build >= EXCHANGE_2013_SP1:
            deleteitem = create_element(
                f'm:{self.SERVICE_NAME}',
                attrs=dict(
                    DeleteType=delete_type,
                    SendMeetingCancellations=send_meeting_cancellations,
                    AffectedTaskOccurrences=affected_task_occurrences,
                    SuppressReadReceipts=suppress_read_receipts,
                )
            )
        else:
            deleteitem = create_element(
                f'm:{self.SERVICE_NAME}',
                attrs=dict(
                    DeleteType=delete_type,
                    SendMeetingCancellations=send_meeting_cancellations,
                    AffectedTaskOccurrences=affected_task_occurrences,
                 )
            )

        item_ids = create_item_ids_element(items=items, version=self.account.version)
        deleteitem.append(item_ids)
        return deleteitem
