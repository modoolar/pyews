# -*- coding: utf-8 -*-
#!/usr/bin/python
##
# Created : Wed Mar 05 11:28:41 IST 2014
##
# Copyright (C) 2014 Sriram Karra <karra.etc@gmail.com>
##
# This file is part of pyews
##
# pyews is free software: you can redistribute it and/or modify it under
# the terms of the GNU Affero General Public License as published by the
# Free Software Foundation, version 3 of the License
##
# pyews is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Affero General Public
# License for more details.
##
# You should have a copy of the license in the doc/ directory of pyews.  If
# not, see <http://www.gnu.org/licenses/>.

import logging
import re
import pdb
import utils
from utils import pretty_xml
from ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from ews.data import DistinguishedFolderId, WellKnownFolderName
from ews.data import FolderClass
from ews.errors import EWSMessageError, EWSCreateFolderError
from ews.errors import EWSDeleteFolderError
from ews.folder import Folder
from ews.contact import Contact

from ews.request_response import GetItemsRequest, GetItemsResponse
from ews.request_response import GetContactsRequest, GetContactsResponse
from ews.request_response import FindItemsRequest, FindItemsResponse
from ews.request_response import CreateItemsRequest, CreateItemsResponse
from ews.request_response import DeleteItemsRequest, DeleteItemsResponse
from ews.request_response import FindItemsLMTRequest, FindItemsLMTResponse
from ews.request_response import (SearchContactByEmailRequest,
                                  SearchContactByEmailResponse)
from ews.request_response import (MoveItemsRequest,
                                  MoveItemsResponse)
from ews.request_response import UpdateItemsRequest, UpdateItemsResponse
from ews.request_response import (SyncFolderItemsRequest,
                                  SyncFolderItemsResponse)
from ews.request_response import (FindCalendarItemsRequest,
                                  FindCalendarItemsResponse)
from ews.request_response import (GetCalendarItemsRequest,
                                  GetCalendarItemsResponse)
from ews.request_response import (FindCalendarItemsRequestBothDate,
                                  FindCalendarItemsResponseBothDate)
from ews.request_response import (FindCalendarItemsRequestDate,
                                  FindCalendarItemsResponseDate)
from ews.request_response import GetAttachmentsRequest
from ews.request_response import CreateAttachmentRequest
from ews.request_response import DeleteAttachmentRequest
from ews.request_response import UpdateCalendarItemsRequest


from tornado import template
from soap import SoapClient, SoapMessageError, QName_T

USER = u''
PWD = u''
EWS_URL = u''

##
# Note: There is a feeeble attemp to mimick the names of classes and methods
# used in the EWS Managed Services API. However the similiarities are merely
# skin-deep, if anything at all.
##


class InvalidUserEmail(Exception):
    pass


class WebCredentials(object):

    def __init__(self, user, pwd, cert=False):
        self.user = user
        self.pwd = pwd
        self.cert = cert


class ExchangeService(object):

    def __init__(self):
        self.ews_ad = None
        self.credentials = None
        self.root_folder = None
        self.loader = template.Loader(utils.REQUESTS_DIR)
        # impersonation mechanisms
        self.primary_smtp_address = None
        self.security_identifier_descriptor = None
        self.principal_name = None

    ##
    # First the methods that are similar to the EWS Managed API.
    # The names might be similar but please note that there is no effort made
    # to really be a complete copy of the Managed API.
    ##
    # FIXME: Each of these results in a EWS Request. We should just have a
    # base Request class along with specific request types derived from the
    # based request that does the needful. Hm. there is no end to this
    # 'properisation'.
    ##

    def AutoDiscoverUrl(self):
        """blame the weird naming on the EWS MS APi."""

        creds = self.credentials
        self.ews_ad = EWSAutoDiscover(creds.user, creds.pwd)
        self.Url = self.ews_ad.discover()

    def CreateFolder(self, parent_id, info):
        """info should be an array of (name, class) tuples. class should be one
        of values in the ews.data.FolderClass enumeration.
        """

        logging.info('Sending folder create request to EWS...')
        req = self._render_template(
            utils.REQ_CREATE_FOLDER,
            parent_folder_id=parent_id,
            folders=info,
            primary_smtp_address=self.primary_smtp_address)
        try:
            resp = self.send(req)
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        logging.info('Sending folder create request to EWS...done')
        return Folder(self, resp)

    def DeleteFolder(self, folder_ids, hard_delete=False):
        """Delete all specified folder ids. If hard_delete is True then the
        folders are completely nuked otherwise they are pushed to the Dumpster
        if that is enabed server side.
        """
        logging.info('pimdb_ex:DeleteFolder() - deleting folder_ids: %s',
                     folder_ids)

        dt = 'HardDelete'if hard_delete else 'SoftDelete'
        req = self._render_template(
            utils.REQ_DELETE_FOLDER,
            delete_type=dt, folder_ids=folder_ids,
            primary_smtp_address=self.primary_smtp_address)
        try:
            resp, node = self.send(req)
            logging.info('pimdb_ex:DeleteFolder() - successfully deleted.')
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        return resp

    def FindCalendarItemsByBothDate(self, folder, start_date, end_date,
                                    eprops_xml=[], ids_only=False):
        """

        """
        req = FindCalendarItemsRequestBothDate(self, folder_id=folder.Id,
                                               start_date=start_date,
                                               end_date=end_date)
        ret = []
        resp = req.execute()
        shells = resp.items
        if shells is not None and len(shells) > 0:
            ret += shells
        if len(ret) > 0 and not ids_only:
            return self.GetCalendarItems([x.itemid for x in ret],
                                         eprops_xml=eprops_xml)
        else:
            return ret

    def FindCalendarItemsByDate(self, folder, start=None, end=None,
                                eprops_xml=[], ids_only=False):
        """

        """
        req = False
        if start is None and end is None:
            return self.FindCalendarItems(folder, eprops_xml=eprops_xml,
                                          ids_only=ids_only)
        elif start and end:
            return self.FindCalendarItemsByBothDate(folder=folder,
                                                    start_date=start,
                                                    end_date=end,
                                                    eprops_xml=eprops_xml,
                                                    ids_only=ids_only)
        else:
            # Either start or end is not None
            i = 0
            ret = []
            while True:
                req = FindCalendarItemsRequestDate(
                    self,
                    batch_size=self.batch_size(),
                    offset=i,
                    folder_id=folder.Id,
                    start_date=start,
                    end_date=end
                )
                resp = req.execute()
                shells = resp.items
                if shells is not None and len(shells) > 0:
                    ret += shells

                if resp.includes_last:
                    break

                i += self.batch_size()
                # just a safety net to avoid inifinite loops
                if i >= folder.TotalCount:
                    break
            if len(ret) > 0 and not ids_only:
                return self.GetCalendarItems([x.itemid for x in ret],
                                             eprops_xml=eprops_xml)
            else:
                return ret

    def FindCalendarItems(self, folder, eprops_xml=[], ids_only=False):
        """

        """
        logging.info('pimdb_ex:FindCalendarItems() - '
                     'fetching items in folder %s...',
                     folder.DisplayName)
        i = 0
        ret = []
        while True:
            req = FindCalendarItemsRequest(self, batch_size=self.batch_size(),
                                           offset=i, folder_id=folder.Id)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            # just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning('pimdb_ex.FindCalendarItems():'
                                'Breaking strange loop')
                break

        logging.info('pimdb_ex:FindCalendarItems() - '
                     'fetching items in folder %s...done',
                     folder.DisplayName)

        if len(ret) > 0 and not ids_only:
            return self.GetCalendarItems([x.itemid for x in ret],
                                         eprops_xml=eprops_xml)
        else:
            return ret

    def FindItems(self, folder, eprops_xml=[], ids_only=False):
        """
        Fetch all the items in the given folder.  folder is an object of type
        ews.folder.Folder. This method will first find all the ItemIds of
        contained items, then go back and fetch all the details for each of
        the items in the folder. We return an array of Item objects of the
        right type.

        eprops_xml is an array of xml representation for additional extended
        properites that need to be fetched

        """

        logging.info('pimdb_ex:FindItems() - fetching items in folder %s...',
                     folder.DisplayName)

        i = 0
        ret = []
        while True:
            req = FindItemsRequest(self, batch_size=self.batch_size(),
                                   offset=i, folder_id=folder.Id)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            # just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning('pimdb_ex.FindItems(): Breaking strange loop')
                break

        logging.info(
            'pimdb_ex:FindItems() - fetching items in folder %s...done',
            folder.DisplayName)

        if len(ret) > 0 and not ids_only:
            return self.GetItems([x.itemid for x in ret],
                                 eprops_xml=eprops_xml)
        else:
            return ret

    def SearchContactByEmail(self, folder, email):
        """
            Search by exact email (case sensitive and fullstring)
            search into EmailAddress1 OR EmailAddress2 OR EmailAddress3
        """
        logging.info(
            'pimdb_ex:SearchContactByEmail() - fetching items in folder %s...',
            folder.DisplayName)
        i = 0
        ret = []
        while True:
            req = SearchContactByEmailRequest(self,
                                              batch_size=self.batch_size(),
                                              offset=i, folder_id=folder.Id,
                                              email=email)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            # just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning(
                    'pimdb_ex.SearchContactByEmail(): Breaking strange loop')
                break

        logging.info(
            'pimdb_ex:SearchContactByEmail() - '
            'fetching items in folder %s...done',
            folder.DisplayName)

        return ret

    def FindItemsLMT(self, folder, lmt):
        """
        Fetch all the items in the given folder that were last modified at or
        after the provided timestamp (lmt). We return an array of Item objects
        of the right type. It will only contain the following fields:

        - ID
        - ChangeKey
        - DisplayName (if it is a contact)
        - Last Modified time

        """

        logging.info(
            'pimdb_ex:FindItemsLMT() - fetching items in folder %s...',
            folder.DisplayName)

        i = 0
        ret = []
        while True:
            req = FindItemsLMTRequest(self, batch_size=self.batch_size(),
                                      offset=i, folder_id=folder.Id, lmt=lmt)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            # just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning(
                    'pimdb_ex.FindItemsLMT(): Breaking strange loop')
                break

        logging.info(
            'pimdb_ex:FindItemsLMT() - fetching items in folder %s...done',
            folder.DisplayName)

        return ret

    def GetCalendarItems(self, calendar_item_ids, eprops_xml=[]):
        logging.info('pimdb_ex:GetCalendarItems() - fetching items....')
        req = GetCalendarItemsRequest(self,
                                      calendar_item_ids=calendar_item_ids,
                                      custom_eprops_xml=eprops_xml)
        resp = req.execute()
        logging.info('pimdb_ex:GetCalendarItems() - fetching items...done')

        return resp.items

    def GetContacts(self, contact_ids, eprops_xml=[]):
        """
        contact_ids is an array of exchange contact ids, and we will fetch that
        stuff and return an array of Item objects.

        @FIXME: Need to make this work in batches to ensure data is not too
        much.
        """

        logging.info('pimdb_ex:GetItems() - fetching items....')
        req = GetContactsRequest(self, contact_ids=contact_ids,
                                 custom_eprops_xml=eprops_xml)
        resp = req.execute()
        logging.info('pimdb_ex:GetItems() - fetching items...done')

        return resp.items

    def GetItems(self, itemids, eprops_xml=[]):
        """
        itemids is an array of itemids, and we will fetch that stuff and
        return an array of Item objects.

        FIXME: Need to make this work in batches to ensure data is not too
        much.
        """

        logging.info('pimdb_ex:GetItems() - fetching items....')
        req = GetItemsRequest(self, itemids=itemids,
                              custom_eprops_xml=eprops_xml)
        resp = req.execute()
        logging.info('pimdb_ex:GetItems() - fetching items...done')

        return resp.items

    def CreateItems(self, folder_id, items):
        """Create items in the exchange store."""
        logging.info('pimdb_ex:CreateItems() - creating items....')
        req = CreateItemsRequest(self, folder_id=folder_id, items=items)
        resp = req.execute()

        logging.info('pimdb_ex:CreateItems() - creating items....done')
        return resp

    def CreateItem(self, folder_id, item):
        """Create one item and return its Id and ChangeKey"""
        logging.info('pimdb_ex:CreateItem() - creating one item....')
        req = CreateItemsRequest(self, folder_id=folder_id, items=[item])
        resp = req.execute()

        logging.info('pimdb_ex:CreateItem() - creating items....done')
        resp_dict = resp.get_itemids()
        # as we only gave one item to create, we can use [0]
        Id = resp_dict.keys()[0]
        CK = resp_dict[Id]
        logging.info('pimdb_ex:CreateItem() - RETURN VALUES: %s %s' % (Id, CK))

        return Id, CK

    def CreateCalendarItem(self,
                           folder_id,
                           item,
                           send_meeting_invitations="SendToNone"):
        logging.info('pimdb_ex:CreateCalendarItem() - creating one item....')
        req = CreateItemsRequest(
            self,
            folder_id=folder_id,
            items=[item],
            send_meeting_invitations=send_meeting_invitations
        )

        resp = req.execute()

        logging.info('pimdb_ex:CreateCalendarItem() - creating items....done')
        resp_dict = resp.get_itemids()
        # as we only gave one item to create, we can use [0]
        Id = resp_dict.keys()[0]
        CK = resp_dict[Id]
        logging.info('pimdb_ex:CreateCalendarItem() '
                     '- RETURN VALUES: %s %s' % (Id, CK))

        return Id, CK

    def CreateAttachment(self, item, attachment):
        req = CreateAttachmentRequest(self, item=item, attachment=attachment)
        return req.execute()

    def DeleteAttachments(self, items):
        req = DeleteAttachmentRequest(self, items=items)
        return req.execute()

    def DeleteCalendarItems(self, itemids,
                            send_meeting_cancellations="SendToNone"):
        logging.info('pimdb_ex:DeleteCalendarItems() - deleting items....')
        req = DeleteItemsRequest(
            self, itemids=itemids,
            calendar=True,
            send_meeting_cancellations=send_meeting_cancellations
        )
        logging.info('pimdb_ex:DeleteCalendarItems() - deleting items....done')
        return req.execute()

    def DeleteItems(self, itemids):
        """Delete items in the exchange store."""

        logging.info('pimdb_ex:DeleteItems() - deleting items....')
        req = DeleteItemsRequest(self, itemids=itemids, calendar=False)
        logging.info('pimdb_ex:DeleteItems() - deleting items....done')

        return req.execute()

    def GetAttachments(self, items):
        """
        Retrieve specified attachment ids
        """
        req = GetAttachmentsRequest(self, items=items)
        resp = req.execute()
        return resp.items

    def MoveItems(self, folder_id, itemids):
        """Move items in the exchange store."""

        logging.info('pimdb_ex:MoveItems() - moveing items....')
        req = MoveItemsRequest(self, folder_id=folder_id, itemids=itemids)
        logging.info('pimdb_ex:MoveItems() - moveing items....done')

        return req.execute()

    def UpdateItems(self, items):
        """
        Fetch updates from the specified folder_id.
        Items in the exchange store.
        """

        logging.info('pimdb_ex:UpdateItems() - updating items....')

        req = UpdateItemsRequest(self, items=items)
        resp = req.execute()

        logging.info('pimdb_ex:UpdateItems() - updating items....done')
        return resp.items

    def UpdateCalendarItems(self, items,
                            send_meeting_invitations="SendToNone"):
        """
        Fetch updates from the specified folder_id.
        Items in the exchange store.
        """

        logging.info('pimdb_ex:UpdateCalendarItems() - updating items....')

        req = UpdateCalendarItemsRequest(
            self,
            items=items,
            send_meeting_invitations=send_meeting_invitations)
        resp = req.execute()

        logging.info('pimdb_ex:UpdateCalendarItems() - updating items....done')
        return resp.items

    def SyncFolderItems(self, folder_id, sync_state):
        """
        Fetch updates from the specified folder_id.
        """

        logging.info('pimdb_ex:SyncFolder() - fetching state...')

        req = SyncFolderItemsRequest(self, folder_id=folder_id,
                                     sync_state=sync_state,
                                     batch_size=self.batch_size())
        resp = req.execute()

        logging.info('pimdb_ex:SyncFolder() - fetching state...done')

        return resp

    ##
    # Some internal messages
    ##

    ##
    # Other external methods
    ##

    def init_soap_client(self):
        self.soap = SoapClient(self.Url,
                               user=self.credentials.user,
                               pwd=self.credentials.pwd,
                               cert=self.credentials.cert)

    def send(self, req, debug=True):
        """
        Will raise a SoapConnectionError if there is a connection problem.
        """

        return self.soap.send(req, debug)

    def get_distinguished_folder(self, name):
        elem = u'<t:DistinguishedFolderId Id="%s"/>' % name
        req = self._render_template(
            utils.REQ_GET_FOLDER,
            folder_ids=elem,
            primary_smtp_address=self.primary_smtp_address)
        return self.soap.send(req)

    def get_root_folder(self):
        if not self.root_folder:
            self.root_folder = Folder.bind(self,
                                           WellKnownFolderName.MsgFolderRoot)
        return self.root_folder

    def batch_size(self):
        return 100

    ##
    # Internal routines
    ##

    def _wsdl_url(self, url=None):
        if not url:
            url = self.Url

        res = re.match('(.*)Exchange.asmx$', url)
        return res.group(1) + 'Services.wsdl'

    # FIXME: To be removed once all the requests become classes
    def _render_template(self, name, **kwargs):
        return self.loader.load(name).generate(**kwargs)

    ##
    # Property getter/setter stuff
    ##

    @property
    def credentials(self):
        return self._credentials

    @credentials.setter
    def credentials(self, c):
        self._credentials = c

    @property
    def Url(self):
        return self._Url

    @Url.setter
    def Url(self, url):
        self._Url = url
        self.wsdl_url = self._wsdl_url()
