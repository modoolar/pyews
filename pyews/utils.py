# -*- coding: utf-8 -*-
import md5
import os
import xml.dom.minidom

CUR_DIR = os.path.dirname(os.path.realpath(__file__))
REQUESTS_DIR = os.path.abspath(os.path.join(CUR_DIR, "templates"))


def template_fn(fn):
    # return os.path.abspath(os.path.join(REQUESTS_DIR, fn))
    return fn

REQ_AUTODIS_EPS = template_fn("autodiscover_endpoints.xml")
REQ_GET_FOLDER = template_fn("get_folder.xml")
REQ_CREATE_FOLDER = template_fn("create_folder.xml")
REQ_DELETE_FOLDER = template_fn("delete_folder.xml")
REQ_BIND_FOLDER = template_fn("bind.xml")
REQ_FIND_FOLDER_ID = template_fn("find_folder_id.xml")
REQ_FIND_FOLDER_DI = template_fn("find_folder_distinguishd.xml")
REQ_FIND_ITEM = template_fn("find_item.xml")
REQ_FIND_ITEM_LMT = template_fn("find_item_lmt.xml")
REQ_FIND_ITEM_MAIL = template_fn("find_contact_by_mail.xml")
REQ_GET_ITEM = template_fn("get_item.xml")
REQ_GET_CONTACT = template_fn("get_contact.xml")
REQ_CREATE_ITEM = template_fn("create_item.xml")
REQ_DELETE_ITEM = template_fn("delete_item.xml")
REQ_UPDATE_ITEM = template_fn("update_item.xml")
REQ_UPDATE_ITEM2 = template_fn("update_item2.xml")
REQ_SYNC_FOLDER = template_fn("sync_folder.xml")
REQ_MOVE_ITEM = template_fn("move_item.xml")
REQ_FIND_CAL_ITEM = template_fn("find_calendar_item.xml")
REQ_GET_CALENDAR_ITEMS = template_fn("get_calendar_item.xml")
REQ_FIND_CAL_ITEM_BOTH_DATE = template_fn(
    "find_calendar_item_by_date_start_end.xml")
REQ_FIND_CAL_ITEM_DATE = template_fn("find_calendar_item_by_date.xml")
REQ_GET_ATTACHMENT = template_fn("get_attachment.xml")
REQ_CREATE_ATT = template_fn("create_attachment.xml")
REQ_DELETE_ATT = template_fn("delete_attachment.xml")


def pretty_xml(x):
    x = xml.dom.minidom.parseString(x).toprettyxml()
    return x


def pretty_eid(x):
    """
    Returns a 32-digit md5 digest of the input. This is useful for printing
    really long and butt-ugly Base64-encoded Entry IDs used by
    Exchange/Outlook
    """

    return md5.new(x).hexdigest()


def safe_int(s):
    """Convert string s into an integer taking into account if s is a hex
    representaiton with a leading 0x."""

    if s[0:2] == '0x':
        return int(s, 16)
    elif s[0:1] == '0':
        return int(s, 8)
    else:
        return int(s, 10)
