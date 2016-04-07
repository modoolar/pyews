
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
from pyews.ews.data import GenderType
from xml.sax.saxutils import escape
_logger = logging.getLogger(__name__)


class CalendarItem(Item):
    def __init__(self, service, parent_fid=None,
                 resp_node=None, mapped_data=None):
        Item.__init__(self, service, parent_fid, resp_node, tag='CalendarItem')
        import pdb; pdb.set_trace()
        if resp_node is not None:
            self._init_from_resp()

    def _init_from_resp(self):
        _logger.info("CalendarItem._init_from_resp")
        pass

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
