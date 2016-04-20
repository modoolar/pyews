
import logging
import sys

from pyews.pyews import WebCredentials, ExchangeService
from pyews.ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from pyews.ews.data import DistinguishedFolderId, WellKnownFolderName
from pyews.ews.data import (FolderClass, GenderType, EmailKey, PhoneKey,
                            PhysicalAddressType, ImAddressType)
from pyews import utils
from pyews.ews.folder import Folder
from pyews.ews.contact import *


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
    cfs = root.FindFolders(types=FolderClass.Contacts)
    # test_fetch_contact_folder(root)

    # odoo_contact_folder_search = cfs[0].FindFolderByDisplayName(
    #     'Odoo Contacts',
    #     types=[FolderClass.Contacts])
    # if odoo_contact_folder_search and len(odoo_contact_folder_search) == 1:
    #     test_delete_folder([odoo_contact_folder_search[0].Id])

    # test_create_folder(cfs[0])

    # test_find_item(cons[0].itemid.text)

    odoo_contact_folder_search = cfs[0].FindFolderByDisplayName(
        'Odoo Contacts',
        types=[FolderClass.Contacts])
    Cid, Cck = test_create_item(ews, odoo_contact_folder_search[0].Id)

    contact = test_find_item(Cid)
    test_update_item(Cid, Cck, contact.parent_fid)
    # cons = test_list_items(root, ids_only=True)

    # iid = 'AAAcAHNrYXJyYUBhc3luay5vbm1pY3Jvc29mdC5jb20ARgAAAAAA6tvK38NMgEiPr'
    #       'dzycecYvAcACf/6iQHYvUyNzrlQXzUQNgAAAAABDwAACf/6iQHYvUyNzrlQXzUQNg'
    #       'AAHykxIwAA'
    # c = test_find_item(iid)
    # test_update_item(iid, c.change_key.value, c.parent_fid)


def bind():
    return Folder.bind(ews, WellKnownFolderName.MsgFolderRoot)


def test_fetch_contact_folder(folder):
    contacts = folder.fetch_all_folders(types=FolderClass.Contacts)
    for f in contacts:
        print 'DisplayName: %s; Id: %s' % (f.DisplayName, f.Id)


def test_create_folder(parent):
    exists = parent.FindFolderByDisplayName(
        "Odoo Contacts",
        types=[FolderClass.Contacts])
    if not exists:
        resp = ews.CreateFolder(parent.Id,
                                [('Odoo Contacts', FolderClass.Contacts)])
    # print utils.pretty_xml(resp)


def test_delete_folder(folders):
    resp = ews.DeleteFolder(folders)


def test_list_items(root, ids_only):
    cfs = root.FindFolders(types=[FolderClass.Contacts])
    contacts = ews.FindItems(cfs[0], ids_only=ids_only)
    print 'Total contacts: ', len(contacts)
    for con in contacts:
        n = con.display_name.value
        print 'Name: %-10s; itemid: %s' % (n, con.itemid)
        print ews.GetItems([con.itemid])

    return contacts


def test_find_item(itemid):
    cons = ews.GetItems([itemid])
    if cons is None or len(cons) <= 0:
        print 'WTF. Could not find itemid ', itemid
    else:
        print cons[0]

    return cons[0]


def test_create_item(ews, fid):
    con = Contact(ews, fid)

    name = "Some Guy"
    function = "Developer"
    phone = "+33 1 23456789"
    mobile = "+33 6 78901234"
    fax = "+33 5 67891234"
    website = "http://www.someguy-website.com"
    email = "someguy@someguy-website.com"
    first_name = "Some"
    last_name = "Guy"
    company_name = "Company."
    file_as = '%s, %s (%s)' % (first_name, last_name.upper(), company_name)
    title_name = "M."

    con.file_as.value = file_as
    con.job_title.value = function
    con.complete_name.given_name.value = first_name
    con.complete_name.surname.value = last_name
    con.complete_name.title.value = title_name
    con.complete_name.middle_name.value = "smg"
    con.company_name.value = company_name
    con.emails.add(EmailKey.Email3, email)
    con.business_home_page.value = website
    con.phones.add(PhoneKey.MobilePhone, mobile)
    con.phones.add(PhoneKey.PrimaryPhone, phone)
    con.ims.add(ImAddressType.Address1, "toto@toto.com")
    con.ims.add(ImAddressType.Address2, "titi@titi.com")

    add1 = PostalAddress()
    add1.add_attrib('Key', PhysicalAddressType.Business)
    add1.street.value = 'business street'
    add1.city.value = 'business city'
    add1.postal_code.value = 'business postal_code'
    add1.country_region.value = 'business country_region'
    con.physical_addresses.add(add1)

    add2 = PostalAddress()
    add2.add_attrib('Key', PhysicalAddressType.Home)
    add2.street.value = 'home street'
    add2.city.value = 'home city'
    add2.postal_code.value = 'home postal_code'
    add2.country_region.value = 'home country_region'
    con.physical_addresses.add(add2)

    add3 = PostalAddress()
    add3.add_attrib('Key', PhysicalAddressType.Other)
    add3.street.value = 'other street'
    add3.city.value = 'other city'
    add3.postal_code.value = 'other postal_code'
    add3.state.value = 'other state'
    add3.country_region.value = 'other country_region'
    con.physical_addresses.add(add3)

    # Main address used in Exchange one in ("Business", "Home", "Other")
    con.postal_address_index.value = PhysicalAddressType.Home

    # con.save()
    Id, CK = ews.CreateItem(fid, con)

    # read created contact
    # read_con = ews.GetItems([Id])
    return Id, CK


def test_update_item(itemid, ck, pfid):
    # con = ews.GetItems([itemid])[0]
    con = Contact(ews, pfid)
    con.itemid.value = itemid
    con.change_key.value = ck
    con.job_title.value = 'Chief comedian'
    con.phones.add(PhoneKey.AssistantPhone, '+91 90088 02194')

    add3 = PostalAddress()
    add3.add_attrib('Key', PhysicalAddressType.Other)
    add3.street.value = 'ororo'
    con.physical_addresses.add(add3)

    ews.UpdateItems([con])

if __name__ == "__main__":
    main()
