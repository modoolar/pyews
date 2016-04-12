
# -*- coding: utf-8 -*-
# Author: Damien Crier
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import logging
import pdb

from item import Item, Field, FieldURI, ExtendedProperty, LastModifiedTime
from pyews.soap import SoapClient, unQName, QName_T
from pyews.utils import pretty_xml
from pyews.ews import mapitags
from pyews.ews.data import MapiPropertyTypeType, MapiPropertyTypeTypeInv
from pyews.ews.data import (LegacyFreeBusyStatusType,
                            CalendarItemTypeType,
                            ConferenceTypeType,
                            )
from pyews.ews.item import (Sensitivity,
                            Importance,
                            ReminderIsSet,
                            Subject,
                            Body,
                            )
from xml.sax.saxutils import escape
from pyews.ews.contact import CField
_logger = logging.getLogger(__name__)


class CalField(Field):

    def __init__(self, tag=None, text=None):
        Field.__init__(self, tag, text)
        self.furi = ('calendar:%s' % tag) if tag else None

    def write_to_xml_update(self):
        _logger.debug("PROCESSED TAG %s : %s" % (self.tag, self.value))
        ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
        s = '<t:FieldURI FieldURI="%s"/>' % self.furi
        s += '\n<t:CalendarItem>'
        s += '\n  <t:%s %s>%s</t:%s>' % (self.tag, ' '.join(ats),
                                         escape(self.value), self.tag)
        s += '\n</t:CalendarItem>'

        return s

    def write_to_xml_delete(self):
        s = '<t:DeleteItemField>'
        s += '<t:FieldURI FieldURI="%s"/>' % self.furi
        s += '</t:DeleteItemField>'
        return s

    def write_to_xml_update2(self):
        return '<t:%s>%s</t:%s>' % (self.tag, escape(self.value), self.tag)


class LegacyFreeBusyStatus(CalField):
    def __init__(self, text=None):
        val_list = LegacyFreeBusyStatusType._props_values()
        err = 'LegacyFreeBusyStatus is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'LegacyFreeBusyStatus', text)


class CalendarItemType(CalField):
    def __init__(self, text=None):
        val_list = CalendarItemTypeType._props_values()
        err = 'CalendarItemType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'CalendarItemType', text)


class ConferenceType(CalField):
    def __init__(self, text=None):
        val_list = ConferenceTypeType._props_values()
        err = 'ConferenceType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'ConferenceType', text)


class CalendarItem(Item):
    def __init__(self, service, parent_fid=None,
                 resp_node=None, mapped_data=None):
        Item.__init__(self, service, parent_fid, resp_node, tag='CalendarItem')

        self.legacy_free_busy_status = LegacyFreeBusyStatus()
        self.calendar_item_type = CalendarItemType()
        self.conference_type = ConferenceType()

        self.tag_property_map.extend([
            (self.legacy_free_busy_status.tag, self.legacy_free_busy_status),
            (self.calendar_item_type.tag, self.calendar_item_type),
            (self.conference_type.tag, self.conference_type),
            ]
        )

        self.mapping_dict_tag_obj = {x: y for x, y in self.tag_property_map}

        # Tags starting with 'Is' are considered as Boolean fields
        # for other fields that must be considered as Boolean ones,
        # add them in the list below
        self.boolean_fields_tag.extend(
            [
            ]
        )

        if resp_node is not None:
            self._init_from_resp()

    def _init_from_resp(self):
        _logger.info("CalendarItem._init_from_resp")
        if self.resp_node is None:
            return

        rnode = self.resp_node
        for child in rnode:
            tag = unQName(child.tag)
            if tag in self.mapping_dict_tag_obj:
                if tag in self.boolean_fields_tag or tag.startswith('Is'):
                    # boolean
                    setattr(self.mapping_dict_tag_obj[tag],
                            'value',
                            child.text == 'true')
                else:
                    setattr(self.mapping_dict_tag_obj[tag],
                            'value',
                            child.text)
            elif tag == 'ExtendedProperty':
                self.add_extended_property(node=child)
            else:
                _logger.warning(
                    'Trying to instanciate an unknown field %s' % tag)

    def add_tagged_property(self, node=None, tag=None, value=None):
        _logger.info("CalendarItem.add_tagged_property")
        pass

# <CalendarItem>
#    <MimeContent/>
#    <ItemId/>
#    <ParentFolderId/>
#    <ItemClass/>
#    <Subject/>
#    <Sensitivity/>
#    <Body/>
#    <Attachments/>
#    <DateTimeReceived/>
#    <Size/>
#    <Categories/>
#    <InReplyTo/>
#    <IsSubmitted/>
#    <IsDraft/>
#    <IsFromMe/>
#    <IsResend/>
#    <IsUnmodified/>
#    <InternetMessageHeaders/>
#    <DateTimeSent/>
#    <DateTimeCreated/>
#    <ResponseObjects/>
#    <ReminderDueBy/>
#    <ReminderIsSet/>
#    <ReminderMinutesBeforeStart/>
#    <DisplayCc/>
#    <DisplayTo/>
#    <HasAttachments/>
#    <ExtendedProperty/>
#    <Culture/>
#    <Start/>
#    <End/>
#    <OriginalStart/>
#    <IsAllDayEvent/>
#    <LegacyFreeBusyStatus/>
#    <Location/>
#    <When/>
#    <IsMeeting/>
#    <IsCancelled/>
#    <IsRecurring/>
#    <MeetingRequestWasSent/>
#    <IsResponseRequested/>
#    <CalendarItemType/>
#    <MyResponseType/>
#    <Organizer/>
#    <RequiredAttendees/>
#    <OptionalAttendees/>
#    <Resources/>
#    <ConflictingMeetingCount/>
#    <AdjacentMeetingCount/>
#    <ConflictingMeetings/>
#    <AdjacentMeetings/>
#    <Duration/>
#    <TimeZone/>
#    <AppointmentReplyTime/>
#    <AppointmentSequenceNumber/>
#    <AppointmentState/>
#    <Recurrence/>
#    <FirstOccurrence/>
#    <LastOccurrence/>
#    <ModifiedOccurrences/>
#    <DeletedOccurrences/>
#    <MeetingTimeZone/>
#    <StartTimeZone/>
#    <EndTimeZone/>
#    <ConferenceType/>
#    <AllowNewTimeProposal/>
#    <IsOnlineMeeting/>
#    <MeetingWorkspaceUrl/>
#    <NetShowUrl/>
#    <EffectiveRights/>
#    <LastModifiedName/>
#    <LastModifiedTime/>
#    <IsAssociated/>
#    <WebClientReadFormQueryString/>
#    <WebClientEditFormQueryString/>
#    <ConversationId/>
#    <UniqueBody/>
# </CalendarItem>

#######################
# Example of reading all properties of a CalendarItem
#######################

# <t:CalendarItem>
#     <t:ItemId ChangeKey="DwAAABY..." Id="AAAQAGMxb2Rvb0B..."/>
#     <t:ParentFolderId ChangeKey="AQAAAA==" Id="AAAQAGMxb2Rvb0BoaWdoY28u..."/>
#     <t:ItemClass>IPM.Appointment</t:ItemClass>
#     <t:Subject>TEST</t:Subject>
#     <t:Sensitivity>Normal</t:Sensitivity>
#     <t:Body BodyType="HTML">&lt;html&gt;
# &lt;head&gt;
# &lt;meta http-equiv=&quot;Content-Type&quot; content=&quot;text/html; charset=utf-8&quot;&gt;
# &lt;meta name=&quot;Generator&quot; content=&quot;Microsoft Exchange Server&quot;&gt;
# &lt;!-- converted from rtf --&gt;
# &lt;style&gt;&lt;!-- .EmailQuote { margin-left: 1pt; padding-left: 4pt; border-left: #800000 2px solid; } --&gt;&lt;/style&gt;
# &lt;/head&gt;
# &lt;body&gt;
# &lt;font face=&quot;Times New Roman&quot; size=&quot;3&quot;&gt;&lt;span style=&quot;font-size:12pt;&quot;&gt;&lt;a name=&quot;BM_BEGIN&quot;&gt;&lt;/a&gt;
# &lt;div&gt;&lt;font face=&quot;Tahoma&quot; size=&quot;2&quot;&gt;&lt;span style=&quot;font-size:10pt;&quot;&gt;Coucou&lt;br&gt;

# &lt;br&gt;

# TEST&lt;br&gt;

# &lt;br&gt;

# --&lt;br&gt;

# test from c1odoo&lt;br&gt;

# &lt;/span&gt;&lt;/font&gt;&lt;/div&gt;
# &lt;/span&gt;&lt;/font&gt;
# &lt;/body&gt;
# &lt;/html&gt;
#     </t:Body>
#     <t:DateTimeReceived>2016-04-07T14:03:53Z</t:DateTimeReceived>
#     <t:Size>2056</t:Size>
#     <t:Importance>Normal</t:Importance>
#     <t:IsSubmitted>false</t:IsSubmitted>
#     <t:IsDraft>false</t:IsDraft>
#     <t:IsFromMe>false</t:IsFromMe>
#     <t:IsResend>false</t:IsResend>
#     <t:IsUnmodified>false</t:IsUnmodified>
#     <t:DateTimeSent>2016-04-07T14:03:53Z</t:DateTimeSent>
#     <t:DateTimeCreated>2016-04-07T14:03:52Z</t:DateTimeCreated>
#     <t:ResponseObjects>
#         <t:ForwardItem/>
#     </t:ResponseObjects>
#     <t:ReminderDueBy>2016-04-07T09:00:00Z</t:ReminderDueBy>
#     <t:ReminderIsSet>true</t:ReminderIsSet>
#     <t:ReminderMinutesBeforeStart>15</t:ReminderMinutesBeforeStart>
#     <t:DisplayCc/>
#     <t:DisplayTo/>
#     <t:HasAttachments>false</t:HasAttachments>
#     <t:Culture>fr-FR</t:Culture>
#     <t:Start>2016-04-07T09:00:00Z</t:Start>
#     <t:End>2016-04-07T10:00:00Z</t:End>
#     <t:IsAllDayEvent>false</t:IsAllDayEvent>
#     <t:LegacyFreeBusyStatus>Busy</t:LegacyFreeBusyStatus>
#     <t:Location>jj</t:Location>
#     <t:IsMeeting>false</t:IsMeeting>
#     <t:IsCancelled>false</t:IsCancelled>
#     <t:IsRecurring>false</t:IsRecurring>
#     <t:MeetingRequestWasSent>false</t:MeetingRequestWasSent>
#     <t:IsResponseRequested>true</t:IsResponseRequested>
#     <t:CalendarItemType>Single</t:CalendarItemType>
#     <t:MyResponseType>Organizer</t:MyResponseType>
#     <t:Organizer>
#         <t:Mailbox>
#             <t:Name>Client 1 Odoo</t:Name>
#             <t:EmailAddress>c1odoo@highco.fr</t:EmailAddress>
#             <t:RoutingType>SMTP</t:RoutingType>
#         </t:Mailbox>
#     </t:Organizer>
#     <t:ConflictingMeetingCount>0</t:ConflictingMeetingCount>
#     <t:AdjacentMeetingCount>0</t:AdjacentMeetingCount>
#     <t:Duration>PT1H</t:Duration>
#     <t:TimeZone>(UTC+01:00) Bruxelles, Copenhague, Madrid, Paris</t:TimeZone>
#     <t:AppointmentSequenceNumber>0</t:AppointmentSequenceNumber>
#     <t:AppointmentState>0</t:AppointmentState>
# </t:CalendarItem>
