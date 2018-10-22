# -*- coding: utf-8 -*-
##
# Created : Thu May 01 11:00:15 IST 2014
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
# s License for more details.
##
# You should have a copy of the license in the doc/ directory of pyews.  If
# not, see <http://www.gnu.org/licenses/>.
import pdb
import logging
import base64
import xml.etree.ElementTree as ET
import pyews.utils as utils

from abc import ABCMeta, abstractmethod
from pyews.soap import SoapClient, QName_S, QName_T, QName_M
from pyews.utils import pretty_xml
from pyews.ews.contact import Contact
from pyews.ews.item import Item, Content
from pyews.ews.calendar import CalendarItem
from pyews.ews.errors import EWSMessageError, EWSResponseError, EWSBaseErrorStr

##
# Base classes
##


class Request(object, metaclass=ABCMeta):
    def __init__(self, ews, template=None):
        self.ews = ews
        self.template = template
        self.kwargs = None
        self.resp = None

    ##
    # Abstract methods
    ##

    @abstractmethod
    def execute(self):
        pass

    ##
    # Public methods
    ##

    def request_server(self, debug=True):
        # pdb.set_trace()
        r = self.ews.loader.load(self.template).generate(**self.kwargs)
        print(r)
        r = utils.pretty_xml(r)
        # uncomment next line to see request sent to exchange server

        if debug:
            logging.debug('Request: %s', r)
        return self.ews.send(r, debug)

    def assert_error(self):
        if self.resp is not None:
            return

        if self.has_errors():
            raise EWSResponseError(self.resp)


class Response(object):

    def __init__(self, req, node):
        self.req = req
        self.node = node
        self.err_cnt = 0
        self.suc_cnt = 0
        self.war_cnt = 0
        self.errors = {}

        # Useful for some response objects.
        self.includes_last = True

        self.parse_for_faults()

    def snarf_includes_last(self):
        gna = SoapClient.get_node_attribute
        last = gna(self.node, 'RootFolder', 'IncludesLastItemInRange')
        self.includes_last = (last == 'true')

        return self.includes_last

    def parse_for_faults(self):
        """
        Check the response xml for any Faults in the XML request. Faults are
        when the server has trouble even  understanding the request.
        """

        self.has_faults = False

        for fault in self.node.iter(QName_S('Fault')):
            self.fault_code = fault.find('faultcode').text
            self.fault_str = fault.find('faultstring').text
            self.has_faults = True
            logging.error('Fault %s found in request : %s' % (self.fault_code,
                                                              self.fault_str))
            raise EWSMessageError(self)

    def parse_for_errors(self, tag, succ_func=None):
        """
        Look in the present response node for all child nodes of given tag,
        while looking for errors and warnings as well.

        Errors are when the server understood our message, but could not do
        what was asked of it, for whatever reason.
        """

        assert self.node is not None

        i = 0
        # pdb.set_trace()
        for gfrm in self.node.iter(tag):
            resp_class = gfrm.attrib['ResponseClass']
            if resp_class == 'Error':
                resp_code = gfrm.find('.//'+QName_M('ResponseCode')).text
                if resp_code == 'ErrorItemNotFound':
                    logging.error(resp_code)
                    continue
                else:
                    self.err_cnt += 1
                    self.errors.update({i: EWSErrorElement(gfrm)})
            elif resp_class == 'Warning':
                self.war_cnt += 1
                # FIXME: Need to handle these
                logging.error('Ignoring a warning response class.')
            else:
                self.suc_cnt += 1
                if succ_func is not None:
                    succ_func(i, gfrm)

            i += 1

        if self.has_errors():
            logging.error('Response.parse: Found %d errors', self.err_cnt)
            for ind, err in list(self.errors.items()):
                logging.error('  Item num %02d - %s', ind, err)

    def has_errors(self):
        return self.err_cnt > 0


class EWSErrorElement(object):

    """
    Wraps an XML response element that represents an erorr response from the
    server
    """

    # <m:GetItemResponseMessage ResponseClass="Error">
    #     <m:MessageText>Id is malformed.</m:MessageText>
    #     <m:ResponseCode>ErrorInvalidIdMalformed</m:ResponseCode>
    #     <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
    #     <m:Items />
    # </m:GetItemResponseMessage>

    def __init__(self, node):
        self.node = node

        t = node.find(QName_M('MessageText'))
        self.msg_text = t.text if t is not None else None

        t = node.find(QName_M('ResponseCode'))
        self.resp_code = t.text if t is not None else None

        t = node.find(QName_M('DescriptiveLinkKey'))
        self.des_link_key = t.text if t is not None else None

    def __str__(self):
        return 'Code: %s; Text: %s' % (self.resp_code, self.msg_text)

##
# Bind
##


class GetFolderRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_BIND_FOLDER)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetFolderResponse(self, self.resp_node)

        return self.resp_obj


class GetFolderResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('GetFolderResponseMessage'))
        for child in self.node.iter(QName_T('Folder')):
            self.folder_node = child
            break

##
# ConvertId
##


class ConvertIdRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_CONVERT_ID)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = ConvertIdResponse(self, self.resp_node)
        return self.resp_obj


class ConvertIdResponse(Response):
    def __init__(self, req, node=None):
        Response.__init__(self, req, node)
        self.items = []
        if node is not None:
                self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.parse_for_errors(
            QName_M('ConvertIdResponseMessage'))

        for cxml in self.node.iter(QName_M('AlternateId')):
            self.items.append(cxml.attrib['Id'])
            break


##
# GetUserConfiguration
##


class UpdateCategoryListRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_UPD_USER_CONF)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = UpdateCategoryListResponse(self, self.resp_node)
        return self.resp_obj


class UpdateCategoryListResponse(Response):
    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
                self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(
            QName_M('UpdateUserConfigurationResponseMessage'))


class GetUserConfigurationRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_USER_CONF)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetUserConfigurationResponse(self, self.resp_node)
        return self.resp_obj


class GetUserConfigurationResponse(Response):
    def __init__(self, req, node=None):
        Response.__init__(self, req, node)
        self.categories = []
        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('GetUserConfigurationResponseMessage'))
        xml_data_b64 = self.node.find('.//'+QName_T('XmlData')).text
        xml_data = base64.b64decode(xml_data_b64)
        self.categories = [a.attrib for a in ET.fromstring(
            xml_data).getchildren()]


##
# CreateItems
##


class CreateAttachmentRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_CREATE_ATT)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = CreateAttachmentResponse(self, self.resp_node)

        return self.resp_obj


class CreateAttachmentResponse(Response):
    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('CreateAttachmentResponseMessage'))


class CreateItemsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_CREATE_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = CreateItemsResponse(self, self.resp_node)

        return self.resp_obj


class CreateItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('CreateItemResponseMessage'),
                              succ_func=self.set_itemids)

    def set_itemids(self, index, node):
        item = self.req.kwargs['items'][index]
        con_node = SoapClient.find_first_child(node, QName_T('ItemId'),
                                               ret='node')
        itemid = con_node.attrib['Id']
        ck = con_node.attrib['ChangeKey']

        item.itemid.set(itemid)
        item.change_key.set(ck)

    def _get_items(self):
        "Returns items provided to in the request to create remote records"

        items = self.req.kwargs['items'] or []
        return items

    def get_itemids(self):
        return {item.itemid.value: item.change_key.value for
                item in self._get_items()}


##
# MoveItems
##

class MoveItemsRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_MOVE_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = MoveItemsResponse(self, self.resp_node)

        return self.resp_obj


class MoveItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('MoveItemResponseMessage'))


##
# DeleteItems
##

class DeleteAttachmentRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_DELETE_ATT)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = DeleteAttachmentResponse(self, self.resp_node)

        return self.resp_obj


class DeleteAttachmentResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('DeleteAttachmentResponseMessage'))


class DeleteItemsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_DELETE_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = DeleteItemsResponse(self, self.resp_node)

        return self.resp_obj


class DeleteItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('DeleteItemResponseMessage'))

##
# FindFolders
##


class FindFoldersRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_FOLDER_ID)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindFoldersResponse(self, self.resp_node)

        return self.resp_obj


class FindFoldersResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        from pyews.ews.folder import Folder as F

        self.parse_for_errors(QName_M('FindFolderResponseMessage'))

        self.folders = []
        for folders in self.node.iter(QName_T('Folders')):
            for child in folders:
                self.folders.append(F(self.req.ews, None, node=child))
            break

##
# FindItems
##


class FindItemsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindItemsResponse(self, self.resp_node)

        return self.resp_obj


class FindItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('Contact')):
            self.items.append(Contact(self, resp_node=cxml))


class FindCalendarItemsRequestBothDate(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_CAL_ITEM_BOTH_DATE)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindCalendarItemsResponseBothDate(self, self.resp_node)

        return self.resp_obj


class FindCalendarItemsResponseBothDate(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('CalendarItem')):
            cal = CalendarItem(self, resp_node=cxml)
            self.items.append(cal)


class FindCalendarItemsRequestDate(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_CAL_ITEM_DATE)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindCalendarItemsResponseDate(self, self.resp_node)

        return self.resp_obj


class FindCalendarItemsResponseDate(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('CalendarItem')):
            cal = CalendarItem(self, resp_node=cxml)
            self.items.append(cal)


class FindCalendarItemsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_CAL_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindCalendarItemsResponse(self, self.resp_node)

        return self.resp_obj


class FindCalendarItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('CalendarItem')):
            cal = CalendarItem(self, resp_node=cxml)
            self.items.append(cal)

##
# SearchContactByEmail
##


class SearchContactByEmailRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_ITEM_MAIL)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        print(('*** WTF: ', self.kwargs))
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = SearchContactByEmailResponse(self, self.resp_node)

        return self.resp_obj


class SearchContactByEmailResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('Contact')):
            self.items.append(Contact(self, resp_node=cxml))


##
# FindItemsLMT
##

class FindItemsLMTRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_ITEM_LMT)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        print(('*** WTF: ', self.kwargs))
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = FindItemsLMTResponse(self, self.resp_node)

        return self.resp_obj


class FindItemsLMTResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.snarf_includes_last()
        self.parse_for_errors(QName_M('FindItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('Contact')):
            self.items.append(Contact(self, resp_node=cxml))


class GetCalendarItemsRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_CALENDAR_ITEMS)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetCalendarItemsResponse(self, self.resp_node)

        return self.resp_obj


class GetCalendarItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def convert(self, ids):
        converted_ids = []
        for item in ids:
            response = ConvertIdRequest(self.req.ews, itemid=item).execute()
            converted_ids.extend(response.items)
        return converted_ids

    def retry_request_with_converted_id(self):
        legacy_ids = self.req.kwargs['calendar_item_ids']
        converted_ids = self.convert(legacy_ids)
        self.req.kwargs.update(calendar_item_ids=converted_ids)
        return self.req.execute()

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('GetItemResponseMessage'))
        errs = []
        if self.has_errors():
            for ind, err in list(self.errors.items()):
                if err.resp_code == 'ErrorInvalidIdMalformedEwsLegacyIdFormat':

                    self.retry_request_with_converted_id()
                else:
                    errs.append(err.msg_text)

        if errs:
            raise EWSBaseErrorStr('\n'.join(errs))

        self.items = []
        xpath = './/%s/%s' % (QName_M('Items'), QName_T('CalendarItem'))
        for cxml in self.node.iterfind(xpath):
            self.items.append(CalendarItem(self.req.ews, resp_node=cxml))

##
# GetContacts
##


class GetContactsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_CONTACT)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetContactsResponse(self, self.resp_node)

        return self.resp_obj


class GetContactsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def convert(self, ids):
        converted_ids = []
        for item in ids:
            response = ConvertIdRequest(self.req.ews, itemid=item).execute()
            converted_ids.extend(response.items)
        return converted_ids

    def retry_request_with_converted_id(self):
        legacy_ids = self.req.kwargs['contact_ids']
        converted_ids = self.convert(legacy_ids)
        self.req.kwargs.update(contact_ids=converted_ids)
        return self.req.execute()

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """
        self.parse_for_errors(QName_M('GetItemResponseMessage'))
        errs = []
        if self.has_errors():
            for ind, err in list(self.errors.items()):
                if err.resp_code == 'ErrorInvalidIdMalformedEwsLegacyIdFormat':

                    self.retry_request_with_converted_id()
                else:
                    errs.append(err.msg_text)

        if errs:
            raise EWSBaseErrorStr('\n'.join(errs))

        self.items = []
        xpath = './/%s/%s' % (QName_M('Items'), QName_T('Contact'))
        for cxml in self.node.iterfind(xpath):
            self.items.append(Contact(self, resp_node=cxml))

##
# GetItems
##


class GetItemsRequest(Request):

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_ITEM)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetContactsResponse(self, self.resp_node)

        return self.resp_obj


# class GetItemsResponse(Response):

#     def __init__(self, req, node=None):
#         Response.__init__(self, req, node)

#         if node is not None:
#             self.init_from_node(node)

#     def init_from_node(self, node):
#         """
#         node is a parsed XML Element containing the response
#         """

#         self.parse_for_errors(QName_M('GetItemResponseMessage'))

#         self.items = []
#         # FIXME: As we support additional item types we will add more such
#         # loops.
#         for cxml in self.node.iter(QName_T('Contact')):
#             self.items.append(Contact(self, resp_node=cxml))


class GetAttachmentsRequest(Request):
    """
    """
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_ATTACHMENT)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

        # self.items_map = {}
        # for item in self.kwargs['items']:
        #     self.items_map[item.itemid.value] = item

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = GetAttachmentsResponse(self, self.resp_node)

        return self.resp_obj


class GetAttachmentsResponse(Response):
    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse_for_errors(QName_M('GetAttachmentResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('Content')):
            cnt = Content(self.req.ews)
            cnt.set(cxml.text)
            self.items.append(cnt)
##
# UpdateItems
##


class UpdateCalendarItemsRequest(Request):
    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_UPDATE_ITEM2)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

        self.items_map = {}
        for item in self.kwargs['items']:
            self.items_map[item.itemid.value] = item

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = UpdateCalendarItemsResponse(self, self.resp_node)
        self.update_change_keys()

        return self.resp_obj

    def update_change_keys(self):
        for resp_item in self.resp_obj.items:
            iid = resp_item.itemid.value
            ck = resp_item.change_key.value

            self.items_map[iid].change_key.set(ck)


class UpdateCalendarItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response. FIXME
        """

        self.parse_for_errors(QName_M('UpdateCalendarItemsResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        xpath = './/%s/%s' % (QName_M('Items'), QName_T('CalendarItem'))
        for cxml in self.node.iterfind(xpath):
            self.items.append(CalendarItem(self.req.ews, resp_node=cxml))


class UpdateItemsRequest(Request):

    """
    Send an update request on the specified items to the server, and save the
    returned changekeys back to the source item objects
    """

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_UPDATE_ITEM2)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

        self.items_map = {}
        for item in self.kwargs['items']:
            self.items_map[item.itemid.value] = item

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = UpdateItemsResponse(self, self.resp_node)
        self.update_change_keys()

        return self.resp_obj

    def update_change_keys(self):
        for resp_item in self.resp_obj.items:
            iid = resp_item.itemid.value
            ck = resp_item.change_key.value

            self.items_map[iid].change_key.set(ck)


class UpdateItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response. FIXME
        """

        self.parse_for_errors(QName_M('UpdateItemResponseMessage'))

        self.items = []
        # FIXME: As we support additional item types we will add more such
        # loops.
        for cxml in self.node.iter(QName_T('Contact')):
            self.items.append(Contact(self, resp_node=cxml))

##
# SyncFolder
##


class SyncFolderItemsRequest(Request):

    """
    Send an update request on the specified items to the server, and save the
    returned changekeys back to the source item objects
    """

    def __init__(self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_SYNC_FOLDER)
        self.kwargs = kwargs
        self.kwargs.update({'primary_smtp_address': ews.primary_smtp_address})

        self.items_map = {}

    ##
    # Implement the abstract methods
    ##

    def execute(self):
        self.resp_node = self.request_server(debug=True)
        self.resp_obj = SyncFolderItemsResponse(self, self.resp_node)

        return self.resp_obj


class SyncFolderItemsResponse(Response):

    def __init__(self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node(self, node):
        """
        node is a parsed XML Element containing the response. FIXME
        """

        find_child = SoapClient.find_first_child

        self.parse_for_errors(QName_M('SyncFolderItemsResponseMessage'))
        self.snarf_includes_last()
        self.sync_state = find_child(node, QName_M('SyncState'))

        self.news = []
        self.mods = []
        self.dels = []

        for create in node.iter(QName_T('Create')):
            for child in create:
                self.news.append(Contact(self.req.ews, resp_node=child))

        for create in node.iter(QName_T('Update')):
            for child in create:
                self.mods.append(Contact(self.req.ews, resp_node=child))

        for create in node.iter(QName_T('Delete')):
            for child in create:
                self.dels.append(Contact(self.req.ews, resp_node=child))
