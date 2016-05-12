# -*- coding: utf-8 -*-
# Author: Damien Crier
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import logging

from item import (ReadOnly,
                  Item,
                  ItemId,
                  ChangeKey,
                  Field,
                  )
from pyews.soap import unQName
from pyews.ews.data import (LegacyFreeBusyStatusType,
                            CalendarItemTypeType,
                            ConferenceTypeType,
                            ResponseTypeType,
                            DaysOfWeekBaseType,
                            DaysOfWeekType,
                            DayOfWeekIndexType,
                            MonthRecurrenceType,
                            )
from xml.sax.saxutils import escape

_logger = logging.getLogger(__name__)


class CalField(Field):

    def __init__(self, tag=None, text=None, boolean=False):
        Field.__init__(self, tag, text)
        self.furi = ('calendar:%s' % tag) if tag else None
        self.boolean = boolean
        self.in_update = True

    def value_as_xml(self):
        if self.value is not None:
            if self.boolean:
                value = escape(unicode(self.value).lower())
            else:
                value = escape(unicode(self.value))
            return value
        return ''

    def write_to_xml_update(self):
        children = self.get_children()

        if (self.in_update and
            ((self.value is not None) or (len(self.attrib) > 0) or
                (len(children) > 0))):

            text = self.value_as_xml()
            ats = self.atts_as_xml()
            cs = self.children_as_xml()

            ret = '<t:SetItemField>'
            ret += '<t:FieldURI FieldURI="%s"/>' % self.furi
            ret += '<t:CalendarItem>'
            ret += '<t:%s %s>%s%s</t:%s>' % (self.tag, ats, text, cs, self.tag)
            ret += '</t:CalendarItem>'
            ret += '</t:SetItemField>'
            return ret
        else:
            return ''

    # def write_to_xml_update(self):
    #     _logger.debug("PROCESSED TAG %s : %s" % (self.tag, self.value))
    #     ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
    #     s = '<t:FieldURI FieldURI="%s"/>' % self.furi
    #     s += '\n<t:CalendarItem>'
    #     s += '\n  <t:%s %s>%s</t:%s>' % (self.tag, ' '.join(ats),
    #                                      escape(self.value), self.tag)
    #     s += '\n</t:CalendarItem>'

    #     return s

    def write_to_xml_delete(self):
        s = '<t:DeleteItemField>'
        s += '<t:FieldURI FieldURI="%s"/>' % self.furi
        s += '</t:DeleteItemField>'
        return s

    # def write_to_xml_update2(self):
    #     return '<t:%s>%s</t:%s>' % (self.tag, escape(self.value), self.tag)


class LegacyFreeBusyStatus(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa566143(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = LegacyFreeBusyStatusType._props_values()
        err = 'LegacyFreeBusyStatus is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'LegacyFreeBusyStatus', text)


class CalendarItemType(ReadOnly, CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa494158(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = CalendarItemTypeType._props_values()
        err = 'CalendarItemType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'CalendarItemType', text)


class ConferenceType(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563529(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = ConferenceTypeType._props_values()
        err = 'ConferenceType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'ConferenceType', text)


class Start(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563943%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'Start', text)


class End(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563714%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'End', text)


class OriginalStart(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580655%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'OriginalStart', text)


class IsAllDayEvent(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa579667(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsAllDayEvent', text, boolean=True)


class Location(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580506%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'Location', text)


class When(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565318(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'When', text)


class IsMeeting(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563544(v=exchg.140).aspx

A value of true indicates that the calendar item is a meeting.
A value of false indicates that the calendar item is an appointment.
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsMeeting', text, boolean=True)
        self.in_update = False


class IsCancelled(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580824(v=exchg.140).aspx

A value of true indicates that the calendar item has been canceled.
A value of false indicates that a calendar item has not been canceled.
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsCancelled', text, boolean=True)
        self.in_update = False


class IsRecurring(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/bb204271(v=exchg.140).aspx

A value of true indicates that the calendar item has been canceled.
A value of false indicates that a calendar item has not been canceled.
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsRecurring', text, boolean=True)
        self.in_update = False


class MeetingRequestWasSent(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564052(v=exchg.140).aspx

A value of true indicates that the calendar item has been canceled.
A value of false indicates that a calendar item has not been canceled.
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'MeetingRequestWasSent', text, boolean=True)
        self.in_update = False


class IsResponseRequested(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563990%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsResponseRequested', text, boolean=True)


class MyResponseType(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564248%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        val_list = ResponseTypeType._props_values()
        err = 'MyResponseType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'MyResponseType', text)
        self.in_update = False


class ResponseType(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565432%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        val_list = ResponseTypeType._props_values()
        err = 'ResponseType is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'ResponseType', text)


class Duration(ReadOnly, CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565298%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'Duration', text)


class TimeZone(ReadOnly, CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564689%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'TimeZone', text)


class AllowNewTimeProposal(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564751%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'AllowNewTimeProposal', text, boolean=True)


class IsOnlineMeeting(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565195(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'IsOnlineMeeting', text, boolean=True)


class MeetingWorkspaceUrl(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563246%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'MeetingWorkspaceUrl', text)


class NetShowUrl(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564533(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'NetShowUrl', text)


class Mailbox(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565036(v=exchg.140).aspx
    """
    class Name(CalField):
        def __init__(self, text=None):
            CalField.__init__(self, 'Name', text)

    class EmailAddress(CalField):
        def __init__(self, text=None):
            CalField.__init__(self, 'EmailAddress', text)

    class RoutingType(CalField):
        def __init__(self, text=None):
            err = 'RoutingType is not valid (only SMTP is valid)'
            assert (text is None or text == 'SMTP'), err
            CalField.__init__(self, 'RoutingType', text)

    def __init__(self, node=None):
        CalField.__init__(self, 'Mailbox')
        self.name = self.Name()
        self.email_address = self.EmailAddress()
        self.routing_type = self.RoutingType('SMTP')
        self.children = [self.name, self.email_address, self.routing_type]

        self.tag_field_mapping = {
            'Name': 'name',
            'EmailAddress': 'email_address',
            'RoutingType': 'routing_type',
        }

        # if node is not None:
        #     self.populate_from_node(node)

    def get_children(self):
        return self.children

    def has_updates(self):
        return True

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class Organizer(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa581279%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'Organizer')
        self.mailbox = Mailbox()

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            if tag == self.mailbox.tag:
                self.mailbox.populate_from_node(child)


class LastResponseTime(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564008%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'LastResponseTime', text)


class Attendee(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580339%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'Attendee')
        self.mailbox = Mailbox()
        self.response_type = ResponseType()
        self.last_response_time = LastResponseTime()

        self.children = [self.mailbox,
                         self.response_type,
                         self.last_response_time
                         ]
        self.to_delete = False

        self.tag_field_mapping = {
            'Mailbox': 'mailbox',
            'ResponseType': 'response_type',
            'LastResponseTime': 'last_response_time',
        }

    def get_children(self):
        return self.children

    def has_updates(self):
        return True

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            if tag == self.mailbox.tag:
                self.mailbox.populate_from_node(child)
            else:
                getattr(self, self.tag_field_mapping[tag]).value = child.text


class RequiredAttendees(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580539%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'RequiredAttendees')
        self.entries = []

    def add(self, att_obj):
        self.entries.append(att_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            att = Attendee()
            att.populate_from_node(child)
            self.entries.append(att)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        s = '<t:SetItemField>'
        s += '\n<t:FieldURI FieldURI="calendar:RequiredAttendees"/>'
        s += '\n<t:CalendarItem>'
        s += '\n  <t:RequiredAttendees>'
        s += self.children_as_xml()
        s += '\n  </t:RequiredAttendees>'
        s += '\n</t:CalendarItem>'
        s += '</t:SetItemField>'
        return s


class OptionalAttendees(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa566002%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'OptionalAttendees')
        self.entries = []

    def add(self, att_obj):
        self.entries.append(att_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            att = Attendee()
            att.populate_from_node(child)
            self.entries.append(att)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        s = '<t:SetItemField>'
        s += '\n<t:FieldURI FieldURI="calendar:OptionalAttendees"/>'
        s += '\n<t:CalendarItem>'
        s += '\n  <t:OptionalAttendees>'
        s += self.children_as_xml()
        s += '\n  </t:OptionalAttendees>'
        s += '\n</t:CalendarItem>'
        s += '</t:SetItemField>'
        return s


class Resources(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564483%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'Resources')
        self.entries = []

    def add(self, att_obj):
        self.entries.append(att_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            att = Attendee()
            att.populate_from_node(child)
            self.entries.append(att)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        s = '<t:SetItemField>'
        s += '\n<t:FieldURI FieldURI="calendar:Resources"/>'
        s += '\n<t:CalendarItem>'
        s += '\n  <t:Resources>'
        s += self.children_as_xml()
        s += '\n  </t:Resources>'
        s += '\n</t:CalendarItem>'
        s += '</t:SetItemField>'
        return s


class AppointmentReplyTime(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563784%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'AppointmentReplyTime', text)


class AppointmentSequenceNumber(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa566064%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'AppointmentSequenceNumber', text)
        self.in_update = False


class AppointmentState(ReadOnly, CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564700%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'AppointmentState', text)


class Occurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565603%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'Occurrence')
        self.itemid = ItemId()
        self.change_key = ChangeKey()
        self.start = Start()
        self.end = End()
        self.original_start = OriginalStart()

        self.tag_field_mapping = {
            'ItemId': 'itemid',
            'Start': 'start',
            'End': 'end',
            'OriginalStart': 'original_start',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            if tag == self.itemid.tag:
                self.itemid.set(child.attrib['Id'])
                self.change_key.set(child.attrib['ChangeKey'])
            else:
                getattr(self, self.tag_field_mapping[tag]).value = child.text


class FirstOccurrence(Occurrence):
    def __init__(self, node=None):
        Occurrence.__init__(self, node)
        self.tag = 'FirstOccurrence'
        self.furi = self.furi.replace('Occurrence', 'FirstOccurrence')


class LastOccurrence(Occurrence):
    def __init__(self, node=None):
        Occurrence.__init__(self, node)
        self.tag = 'LastOccurrence'
        self.furi = self.furi.replace('Occurrence', 'LastOccurrence')


class ModifiedOccurrences(CalField):
    def __init__(self, node=None):
        CalField.__init__(self, 'ModifiedOccurrences')
        self.entries = []

    def add(self, occ_obj):
        self.entries.append(occ_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            occ = Occurrence()
            occ.populate_from_node(child)
            self.entries.append(occ)


class DeletedOccurrence(Occurrence):
    def __init__(self, node=None):
        Occurrence.__init__(self, node)
        self.tag = 'DeletedOccurrence'
        self.furi = self.furi.replace('Occurrence', 'DeletedOccurrence')


class DeletedOccurrences(CalField):
    def __init__(self, node=None):
        CalField.__init__(self, 'DeletedOccurrences')
        self.entries = []

    def add(self, occ_obj):
        self.entries.append(occ_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            occ = DeletedOccurrence()
            occ.populate_from_node(child)
            self.entries.append(occ)


class Interval(CalField):
    def __init__(self, text=None):
        CalField.__init__(self, 'Interval', text)


class DaysOfWeek(CalField):
    def __init__(self, text=None):
        vals = DaysOfWeekType._props_values()
        err = 'DaysOfWeek is not in the list %s' % vals
        assert (text is None or all([x in vals for x in text.split()])), err
        CalField.__init__(self, 'DaysOfWeek', text)


class DayOfWeekIndex(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa581350(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = DayOfWeekIndexType._props_values()
        err = 'DayOfWeekIndex is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'DayOfWeekIndex', text)


class DayOfMonth(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563211%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        err = 'DayOfMonth is must be between 1 and 31'
        if text is not None:
            try:
                int_text = int(text)
                assert (text is None or int_text in range(1, 32)), err
                CalField.__init__(self, 'DayOfMonth', text)
            except ValueError:
                _logger.error('Unable to convert %s to int type.'
                              'Considering value as None.' % text)
                CalField.__init__(self, 'DayOfMonth', None)
        else:
            CalField.__init__(self, 'DayOfMonth')


class FirstDayOfWeek(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/ff709517(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = DaysOfWeekBaseType._props_values()
        err = 'FirstDayOfWeek is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'FirstDayOfWeek', text)


class MonthRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565486%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        val_list = MonthRecurrenceType._props_values()
        err = 'Month is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        CalField.__init__(self, 'Month', text)


class RelativeYearlyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/bb204113(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'RelativeYearlyRecurrence')
        self.days_of_week = DaysOfWeek()
        self.day_of_week_index = DayOfWeekIndex()
        self.month = MonthRecurrence()

        self.children = [self.days_of_week, self.day_of_week_index, self.month]

        self.tag_field_mapping = {
            'DaysOfWeek': 'days_of_week',
            'DayOfWeekIndex': 'day_of_week_index',
            'Month': 'month',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class AbsoluteYearlyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564242(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'AbsoluteYearlyRecurrence')
        self.day_of_month = DayOfMonth()
        self.month = MonthRecurrence()

        self.children = [self.day_of_month, self.month]

        self.tag_field_mapping = {
            'DayOfMonth': 'day_of_month',
            'Month': 'month',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class RelativeMonthlyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564558(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'RelativeMonthlyRecurrence')
        self.days_of_week = DaysOfWeek()
        self.day_of_week_index = DayOfWeekIndex()
        self.interval = Interval()

        self.children = [self.interval,
                         self.days_of_week,
                         self.day_of_week_index]

        self.tag_field_mapping = {
            'DaysOfWeek': 'days_of_week',
            'DayOfWeekIndex': 'day_of_week_index',
            'Interval': 'interval',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class AbsoluteMonthlyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa493844(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'AbsoluteMonthlyRecurrence')
        self.day_of_month = DayOfMonth()
        self.interval = Interval()

        self.children = [self.interval, self.day_of_month]

        self.tag_field_mapping = {
            'DayOfMonth': 'day_of_month',
            'Interval': 'interval',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class WeeklyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563500(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'WeeklyRecurrence')
        self.days_of_week = DaysOfWeek()
        self.first_day_of_week = FirstDayOfWeek()
        self.interval = Interval()

        self.children = [self.interval,
                         self.days_of_week,
                         self.first_day_of_week]

        self.tag_field_mapping = {
            'DaysOfWeek': 'days_of_week',
            'FirstDayOfWeek': 'first_day_of_week',
            'Interval': 'interval',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class DailyRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563228(v=exchg.140).aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'DailyRecurrence')
        self.interval = Interval()

        self.children = [self.interval]

        self.tag_field_mapping = {
            'Interval': 'interval',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class DateRecurrence(CalField):
    """
Not to be used directly
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'DateRecurrence', text)


class StartDateRecurrence(DateRecurrence):
    """
https://msdn.microsoft.com/en-us/library/office/aa565019%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        DateRecurrence.__init__(self, text)
        self.tag = 'StartDate'
        self.furi = self.furi.replace('DateRecurrence', 'StartDate')


class EndDateRecurrence(DateRecurrence):
    """
https://msdn.microsoft.com/en-us/library/office/aa493828%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        DateRecurrence.__init__(self, text)
        self.tag = 'EndDate'
        self.furi = self.furi.replace('DateRecurrence', 'EndDate')


class NumberOfOccurrences(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564329%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'NumberOfOccurrences', text)


class NoEndRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564699%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'NoEndRecurrence')
        self.start_date = StartDateRecurrence()

        self.children = [self.start_date]

        self.tag_field_mapping = {
            'StartDate': 'start_date',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class EndRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa564536%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'EndDateRecurrence')
        self.start_date = StartDateRecurrence()
        self.end_date = EndDateRecurrence()

        self.children = [self.start_date, self.end_date]

        self.tag_field_mapping = {
            'StartDate': 'start_date',
            'EndDate': 'end_date',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class NumberedRecurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580960%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'NumberedRecurrence')
        self.start_date = StartDateRecurrence()
        self.nb_occurrences = NumberOfOccurrences()

        self.children = [self.start_date, self.nb_occurrences]

        self.tag_field_mapping = {
            'StartDate': 'start_date',
            'NumberOfOccurrences': 'nb_occurrences',
        }

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text


class Recurrence(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580471%28v=exchg.140%29.aspx
    """
    def __init__(self, node=None):
        CalField.__init__(self, 'Recurrence')

        self.rel_year_rec = RelativeYearlyRecurrence()
        self.abs_year_rec = AbsoluteYearlyRecurrence()
        self.rel_month_rec = RelativeMonthlyRecurrence()
        self.abs_month_rec = AbsoluteMonthlyRecurrence()
        self.week_rec = WeeklyRecurrence()
        self.day_rec = DailyRecurrence()

        self.no_end_rec = NoEndRecurrence()
        self.end_date_rec = EndRecurrence()
        self.numbered_rec = NumberedRecurrence()

        # self.children = [
        #     self.rel_year_rec,
        #     self.abs_year_rec,
        #     self.rel_month_rec,
        #     self.abs_month_rec,
        #     self.week_rec,
        #     self.day_rec,
        #     self.no_end_rec,
        #     self.end_date_rec,
        #     self.numbered_rec,
        # ]

        self.tag_field_mapping = {
            'RelativeYearlyRecurrence': 'rel_year_rec',
            'AbsoluteYearlyRecurrence': 'abs_year_rec',
            'RelativeMonthlyRecurrence': 'rel_month_rec',
            'AbsoluteMonthlyRecurrence': 'abs_month_rec',
            'WeeklyRecurrence': 'week_rec',
            'DailyRecurrence': 'day_rec',
            'NoEndRecurrence': 'no_end_rec',
            'EndDateRecurrence': 'end_date_rec',
            'NumberedRecurrence': 'numbered_rec',
        }

        self.children = self.get_children()

    def write_to_xml_update(self):
        s = '<t:SetItemField>'
        s += '\n<t:FieldURI FieldURI="calendar:Recurrence"/>'
        s += '\n<t:CalendarItem>'
        s += '\n  <t:Recurrence>'
        s += self.children_as_xml()
        s += '\n  </t:Recurrence>'
        s += '\n</t:CalendarItem>'
        s += '</t:SetItemField>'
        return s

    def has_updates(self):
        return True

    def get_children(self):
        """
        """
        FIELD_LIST = self.tag_field_mapping.values()

        end_type_fields = ['no_end_rec', 'end_date_rec', 'numbered_rec']
        end_type_list = [getattr(self, x).start_date.value is None
                         for x in end_type_fields]

        rec_type1_fields = ['rel_year_rec', 'abs_year_rec']
        rec_type1_list = [getattr(self, x).month.value is None
                          for x in rec_type1_fields]
        rec_type_fields = (
            list(set(FIELD_LIST) -
                 set(end_type_fields) -
                 set(rec_type1_fields))
        )
        rec_type_list = [getattr(self, x).interval.value is None
                         for x in rec_type_fields]

        if all(rec_type_list + rec_type1_list + end_type_list):
            return []

        return self._check_instance_pattern()

    def _check_instance_pattern(self):
        """
        check that instance properties filled
        match a valid pattern for Exchange
        """
        rec_type = self._check_recurrence_type()
        end_type = self._check_end_type()

        return getattr(self, rec_type), getattr(self, end_type)

    def _check_recurrence_type(self):
        """
        check only one filled over
            'RelativeYearlyRecurrence'
            'AbsoluteYearlyRecurrence'
            'RelativeMonthlyRecurrence'
            'AbsoluteMonthlyRecurrence'
            'WeeklyRecurrence'
            'DailyRecurrence'
        """
        msg = 'Recurrence end type error.\n'
        msg += 'You have to assign exactly one of the following properties:\n'
        msg += 'RelativeYearlyRecurrence, AbsoluteYearlyRecurrence, '
        msg += 'RelativeMonthlyRecurrence, AbsoluteMonthlyRecurrence, '
        msg += 'WeeklyRecurrence, DailyRecurrence'

        aaa = [(x, getattr(getattr(self, x), y).value) for x, y in (
            ('rel_year_rec', 'month'),
            ('abs_year_rec', 'month'),
            ('rel_month_rec', 'interval'),
            ('abs_month_rec', 'interval'),
            ('week_rec', 'interval'),
            ('day_rec', 'interval'))
        ]

        none_nb = [x[1] for x in aaa].count(None)
        if none_nb < 1:
            # not OK
            raise ValueError(msg)
        elif none_nb == 5:
            # OK
            # [0] because in this case, there will be always one
            rec_type = [x[0] for x in aaa if x[1] is not None][0]
            return rec_type
        elif none_nb > 5:
            # not OK
            raise ValueError(msg)

    def _check_end_type(self):
        """
        check only one filled over
            'NoEndRecurrence'
            'EndDateRecurrence'
            'NumberedRecurrence'
        """
        msg = 'Recurrence end type error.\n'
        msg += 'You have to assign exactly one of the following properties:\n'
        msg += 'NoEndRecurrence, EndDateRecurrence, NumberedRecurrence'

        aaa = [(x, getattr(self, x).start_date.value) for x in (
            'no_end_rec',
            'end_date_rec',
            'numbered_rec')
        ]
        none_nb = [x[1] for x in aaa].count(None)
        if none_nb < 2:
            # not OK
            raise ValueError(msg)
        elif none_nb == 2:
            # OK
            # [0] because in this case, there will be always one
            end_type = [x[0] for x in aaa if x[1] is not None][0]
            return end_type
        elif none_nb > 2:
            # not OK
            raise ValueError(msg)

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self,
                    self.tag_field_mapping[tag]).populate_from_node(child)


class AdjacentMeetingCount(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580264(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'AdjacentMeetingCount', text)
        self.in_update = False


class ConflictingMeetingCount(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa563429(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        CalField.__init__(self, 'ConflictingMeetingCount', text)
        self.in_update = False


class ConflictingMeetings(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa565461%28v=exchg.140%29.aspx
    """
    def __init__(self, service, node=None):
        CalField.__init__(self, 'ConflictingMeetings')
        self.service = service
        self.entries = []
        self.in_update = False

    def add(self, cal_obj):
        self.entries.append(cal_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            cal = CalendarItem(self.service, resp_node=child)
            self.entries.append(cal)


class AdjacentMeetings(CalField):
    """
https://msdn.microsoft.com/en-us/library/office/aa580822%28v=exchg.140%29.aspx
    """
    def __init__(self, service, node=None):
        CalField.__init__(self, 'AdjacentMeetings')
        self.service = service
        self.entries = []
        self.in_update = False

    def add(self, cal_obj):
        self.entries.append(cal_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            cal = CalendarItem(self.service, resp_node=child)
            self.entries.append(cal)


class CalendarItem(Item):
    """
https://msdn.microsoft.com/en-us/library/office/aa564765(v=exchg.140).aspx
    """
    def __init__(self, service, parent_fid=None,
                 resp_node=None, mapped_data=None):
        self.legacy_free_busy_status = LegacyFreeBusyStatus()
        self.calendar_item_type = CalendarItemType()
        self.conference_type = ConferenceType()
        self.start = Start()
        self.end = End()
        self.original_start = OriginalStart()
        self.is_all_day_event = IsAllDayEvent()
        self.location = Location()
        self.when = When()
        self.is_meeting = IsMeeting()
        self.is_cancelled = IsCancelled()
        self.is_recurring = IsRecurring()
        self.meeting_request_was_sent = MeetingRequestWasSent()
        self.is_response_requested = IsResponseRequested()
        self.my_response_type = MyResponseType()
        self.duration = Duration()
        self.timezone = TimeZone()

        self.organizer = Organizer()
        self.required_attendees = RequiredAttendees()
        self.optional_attendees = OptionalAttendees()
        self.resources = Resources()

        self.appointment_reply_time = AppointmentReplyTime()
        self.appointment_seq_number = AppointmentSequenceNumber()
        self.appointment_state = AppointmentState()

        self.first_occurrence = FirstOccurrence()
        self.last_occurrence = LastOccurrence()
        self.modified_occurrences = ModifiedOccurrences()
        self.deleted_occurrences = DeletedOccurrences()
        self.recurrence = Recurrence()

        self.allow_new_time_proposal = AllowNewTimeProposal()
        self.is_online_meeting = IsOnlineMeeting()
        self.meeting_workspace_url = MeetingWorkspaceUrl()
        self.netshow_url = NetShowUrl()

        self.adjacent_meeting_count = AdjacentMeetingCount()
        self.conflicting_meeting_count = ConflictingMeetingCount()
        self.adjacent_meetings = AdjacentMeetings(service)
        self.conflicting_meetings = ConflictingMeetings(service)

        # /!\/!\/!\
        # Ordered list !
        # Fields must be in this order for the xml request to be valid
        # /!\/!\/!\
        self.calendar_tag_property_map = [
            (self.start.tag, self.start),
            (self.end.tag, self.end),
            (self.original_start.tag, self.original_start),
            (self.is_all_day_event.tag, self.is_all_day_event),
            (self.legacy_free_busy_status.tag, self.legacy_free_busy_status),
            (self.location.tag, self.location),
            (self.when.tag, self.when),
            (self.is_meeting.tag, self.is_meeting),
            (self.is_cancelled.tag, self.is_cancelled),
            (self.is_recurring.tag, self.is_recurring),
            (self.meeting_request_was_sent.tag, self.meeting_request_was_sent),
            (self.is_response_requested.tag, self.is_response_requested),
            (self.calendar_item_type.tag, self.calendar_item_type),
            (self.my_response_type.tag, self.my_response_type),
            (self.organizer.tag, self.organizer),
            (self.required_attendees.tag, self.required_attendees),
            (self.optional_attendees.tag, self.optional_attendees),
            (self.resources.tag, self.resources),
            (self.conflicting_meeting_count.tag,
             self.conflicting_meeting_count),
            (self.adjacent_meeting_count.tag, self.adjacent_meeting_count),
            (self.conflicting_meetings.tag, self.conflicting_meetings),
            (self.adjacent_meetings.tag, self.adjacent_meetings),
            (self.duration.tag, self.duration),
            (self.timezone.tag, self.timezone),
            (self.appointment_reply_time.tag, self.appointment_reply_time),
            (self.appointment_seq_number.tag, self.appointment_seq_number),
            (self.appointment_state.tag, self.appointment_state),
            (self.recurrence.tag, self.recurrence),
            (self.first_occurrence.tag, self.first_occurrence),
            (self.last_occurrence.tag, self.last_occurrence),
            (self.modified_occurrences.tag, self.modified_occurrences),
            (self.deleted_occurrences.tag, self.deleted_occurrences),
            (self.conference_type.tag, self.conference_type),
            (self.allow_new_time_proposal.tag, self.allow_new_time_proposal),
            (self.is_online_meeting.tag, self.is_online_meeting),
            (self.meeting_workspace_url.tag, self.meeting_workspace_url),
            (self.netshow_url.tag, self.netshow_url),
            ]

        # Tags starting with 'Is' are considered as Boolean fields
        # for other fields that must be considered as Boolean ones,
        # add them in the list below
        boolean_fields_tag = [
            self.meeting_request_was_sent.tag,
            self.allow_new_time_proposal.tag,
        ]

        Item.__init__(self, service, parent_fid, resp_node, tag='CalendarItem',
                      additional_properties=self.calendar_tag_property_map,
                      additional_boolean_fields_tag=boolean_fields_tag)

    def _process_tag(self, child, tag):
        res = super(CalendarItem, self)._process_tag(child, tag)

        if not res:
            if tag == self.organizer.tag:
                self.organizer.populate_from_node(child)
            elif tag == self.required_attendees.tag:
                self.required_attendees.populate_from_node(child)
            elif tag == self.optional_attendees.tag:
                self.optional_attendees.populate_from_node(child)
            elif tag == self.resources.tag:
                self.resources.populate_from_node(child)
            elif tag == self.first_occurrence.tag:
                self.first_occurrence.populate_from_node(child)
            elif tag == self.last_occurrence.tag:
                self.last_occurrence.populate_from_node(child)
            elif tag == self.modified_occurrences.tag:
                self.modified_occurrences.populate_from_node(child)
            elif tag == self.deleted_occurrences.tag:
                self.deleted_occurrences.populate_from_node(child)
            elif tag == self.recurrence.tag:
                self.recurrence.populate_from_node(child)
            elif tag == self.adjacent_meetings.tag:
                self.adjacent_meetings.populate_from_node(child)
            elif tag == self.conflicting_meetings.tag:
                self.conflicting_meetings.populate_from_node(child)
            else:
                return False
        return True

    def add_tagged_property(self, node=None, tag=None, value=None):
        _logger.info("CalendarItem.add_tagged_property")
        pass

    def get_children(self):
        # The order of these fields is critical. I know, it's crazy.
        self.children = [
            self.subject,
            self.sensitivity,
            self.body,
            self.attachments,
            self.categories,
            self.reminder_due_by,
            self.is_reminder_set,
            self.reminder_minutes_before_start,
        ] + [x[1] for x in self.calendar_tag_property_map]

        return self.children

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
