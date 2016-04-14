# -*- coding: utf-8 -*-
# Author: Damien Crier
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import logging
import sys

from pyews.pyews import WebCredentials, ExchangeService
from pyews.ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from pyews.ews.data import DistinguishedFolderId, WellKnownFolderName
from pyews.ews.data import (FolderClass, GenderType, EmailKey, PhoneKey,
                            PhysicalAddressType, ImAddressType)
from pyews import utils
from pyews.ews.folder import Folder
from pyews.ews.calendar import *


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
    ews.primary_smtp_address = "c1odoo@linkingup.org"

    try:
        ews.AutoDiscoverUrl()
    except ExchangeAutoDiscoverError as e:
        logging.info('ExchangeAutoDiscoverError: %s', e)
        logging.info('Falling back on manual url setting.')
        ews.Url = EWS_URL

    ews.init_soap_client()

    root = bind()
    cfs = root.FindFolders(types=FolderClass.Calendars)
    cals = ews.FindCalendarItems(cfs[0])
    start_date = '2016-04-06T00:00:00Z'
    end_date = '2016-04-08T00:00:00Z'
    find_cals = ews.FindCalendarItemsByBothDate(cfs[0],
                                                start_date=start_date,
                                                end_date=end_date)
    find_cals2 = ews.FindCalendarItemsByDate(cfs[0],
                                             start=end_date)
    find_cals3 = ews.FindCalendarItemsByDate(cfs[0],
                                             end=start_date)
    find_cals4 = ews.FindCalendarItemsByDate(cfs[0])
    assert len(cals) == len(find_cals4)

    cal_id, cal_ck = test_create_calendar_item(ews, cfs[0].Id)


def test_create_calendar_item(ews, fid):
    cal = CalendarItem(ews, fid)

    subject = "TEST from script"
    start_date = "2016-04-14T13:00:00Z"
    end_date = "2016-04-14T17:00:00Z"

    cal.subject.value = subject
    cal.start.value = start_date
    cal.end.value = end_date


    import pdb; pdb.set_trace()
    Id, CK = ews.CreateItem(fid, cal)

    # read created contact
    # read_con = ews.GetItems([Id])
    return Id, CK

def bind():
    return Folder.bind(ews, WellKnownFolderName.MsgFolderRoot)

if __name__ == "__main__":
    main()
