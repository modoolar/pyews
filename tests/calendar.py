# -*- coding: utf-8 -*-
# Author: Damien Crier
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import logging
import sys
import base64

from pyews.pyews import WebCredentials, ExchangeService
from pyews.ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from pyews.ews.data import *
from pyews import utils
from pyews.ews.folder import Folder
from pyews.ews.calendar import *
from pyews.ews.item import *


def main():
    logging.getLogger().setLevel(logging.DEBUG)

    global USER, PWD, EWS_URL, ews

    with open('auth.txt', 'r') as inf:
        USER = inf.readline().strip()
        PWD = inf.readline().strip()
        EWS_URL = inf.readline().strip()

        logging.debug('Username: %s; Url: %s', USER, EWS_URL)

    creds = WebCredentials(USER, PWD)
    ews = ExchangeService()
    ews.credentials = creds
    ews.primary_smtp_address = "c2odoo@linkingup.org"

    try:
        ews.AutoDiscoverUrl()
    except ExchangeAutoDiscoverError as e:
        logging.info('ExchangeAutoDiscoverError: %s', e)
        logging.info('Falling back on manual url setting.')
        ews.Url = EWS_URL

    ews.init_soap_client()

    root = bind()
    cfs = root.FindFolders(types=FolderClass.Calendars)
    # cals = ews.FindCalendarItems(cfs[0])

    # att_id = 'AAAQAGMxb2Rvb0BoaWdoY28uZnIARgAAAAAAOO3rSaghRkyF3KhRUnXPkwcAAIlRgN5j1Emsqr28n7FzvwAAAHTcqAAAAIlRgN5j1Emsqr28n7FzvwAABPR66AAAARIAEACA04EHnuGwR5js5uow3i+q'
    # import pdb; pdb.set_trace()
    # toto = ews.GetAttachments([att_id])

    start_date = '2016-05-16T00:00:00Z'
    end_date = '2016-05-20T00:00:00Z'
    find_cals = ews.FindCalendarItemsByBothDate(cfs[1],
                                                start_date=start_date,
                                                end_date=end_date)
    import pdb; pdb.set_trace()
    # find_cals2 = ews.FindCalendarItemsByDate(cfs[0],
    #                                          start=end_date)
    # find_cals3 = ews.FindCalendarItemsByDate(cfs[0],
    #                                          end=start_date)
    # find_cals4 = ews.FindCalendarItemsByDate(cfs[0])
    # assert len(cals) == len(find_cals4)

    # cal_id, cal_ck = test_create_calendar_item(ews, cfs[0].Id)

    # cal_id = "AAAQAGMxb2Rvb0BoaWdoY28uZnIARgAAAAAAOO3rSaghRkyF3KhRUnXPkwcAAIlRgN5j1Emsqr28n7FzvwAAAHTcqAAAAIlRgN5j1Emsqr28n7FzvwAABPR66AAA"
    # cal_ck = "DwAAABYAAAAAiVGA3mPUSayqvbyfsXO/AAAE9Hv4"

    # test = ews.GetCalendarItems([cal_id])
    # test[0].subject.value = "JJJD BLABLA"
    # ews.UpdateCalendarItems(test)

    # # test_create_attachment(ews, test[0])
    # test_delete_attachment(ews, test[0], 'TEST.TXT')


def test_delete_attachment(ews, item, fname):
    new_read = ews.GetCalendarItems([item.itemid.value])
    import pdb; pdb.set_trace()
    for attachment in new_read[0].attachments.entries:
        if attachment.name.value == fname:
            ews.DeleteAttachments([attachment.attachment_id.value])


def test_create_attachment(ews, item):
    attach = FileAttachment(ews)
    attach.name.set('TEST.TXT')
    file_content = 'Coucou. test creation PJ from Python'
    attach.content.set(base64.b64encode(file_content))
    # item.attachments.add(attach)
    ews.CreateAttachment(item.itemid.value, attach)


def test_create_calendar_item(ews, fid):
    cal = CalendarItem(ews, fid)

    subject = "TEST from script"
    start_date = "2016-04-14T13:00:00Z"
    end_date = "2016-04-14T17:00:00Z"

    cal.subject.value = subject
    cal.start.value = start_date
    cal.end.value = end_date
    cal.calendar_item_type.value = CalendarItemTypeType.Single
    cal.legacy_free_busy_status = LegacyFreeBusyStatusType.Busy

    # attendees
    att1 = Attendee()
    att1.response_type.value = ResponseTypeType.Tentative
    att1.mailbox.name.value = 'TOTO'
    att1.mailbox.email_address.value = 'TOTO@toto.com'
    cal.required_attendees.add(att1)

    # recurrence
    # cal.recurrence.week_rec.interval.value = 1
    # cal.recurrence.week_rec.days_of_week.value = DaysOfWeekBaseType.Thursday
    # cal.recurrence.no_end_rec.start_date.value = start_date[:10]

    Id, CK = ews.CreateCalendarItem(fid, cal)

    # read created contact
    # read_con = ews.GetItems([Id])
    return Id, CK


def bind():
    return Folder.bind(ews, WellKnownFolderName.MsgFolderRoot)

if __name__ == "__main__":
    main()
