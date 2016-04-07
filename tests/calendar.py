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
    import pdb; pdb.set_trace()



def bind():
    return Folder.bind(ews, WellKnownFolderName.MsgFolderRoot)

if __name__ == "__main__":
    main()
