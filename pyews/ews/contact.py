# -*- coding: utf-8 -*-
##
# Created : Wed Apr 02 18:02:51 IST 2014
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
import pdb

from item import (Item,
                  Field,
                  FieldURI,
                  ExtendedProperty,
                  LastModifiedTime,
                  Categories,
                  )
from pyews.soap import SoapClient, unQName, QName_T
from pyews.utils import pretty_xml
from pyews.ews import mapitags
from pyews.ews.data import MapiPropertyTypeType, MapiPropertyTypeTypeInv
from pyews.ews.data import GenderType
from xml.sax.saxutils import escape
_logger = logging.getLogger(__name__)


class CField(Field):

    def __init__(self, tag=None, text=None):
        Field.__init__(self, tag, text)
        self.furi = ('contacts:%s' % tag) if tag else None

    def write_to_xml_update(self):
        _logger.debug("PROCESSED TAG %s : %s" % (self.tag, self.value))
        ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
        s = '<t:SetItemField>'
        s += '<t:FieldURI FieldURI="%s"/>' % self.furi
        s += '\n<t:Contact>'
        s += '\n  <t:%s %s>%s</t:%s>' % (self.tag, ' '.join(ats),
                                         escape(self.value), self.tag)
        s += '\n</t:Contact>'
        s += '</t:SetItemField>'

        return s

    def write_to_xml_delete(self):
        s = '<t:DeleteItemField>'
        s += '<t:FieldURI FieldURI="%s"/>' % self.furi
        s += '</t:DeleteItemField>'
        return s

    def write_to_xml_update2(self):
        return '<t:%s>%s</t:%s>' % (self.tag, escape(self.value), self.tag)


class FileAs(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'FileAs', text)


class FileAsMapping(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'FileAsMapping', text)


class Alias(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Alias', text)


class DisplayName(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'DisplayName', text)


class CompleteName(CField):
    ##
    # Exchnage handling of some of the name fields can, at best, be described
    # as strange.... Some of these fields can be found outside of this
    # complex data type. However we will try to make it simple and only use
    # CompelteName field when possible
    ##

    class Title(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Title', text)

    class FirstName(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'FirstName', text)

    class GivenName(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'GivenName', text)

    class MiddleName(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'MiddleName', text)

    class LastName(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'LastName', text)

    class Surname(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Surname', text)

    class Suffix(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Suffix', text)

    class Initials(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Initials', text)

    class Nickname(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Nickname', text)

    class FullName(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'FullName', text)

    def __init__(self, text=None):
        CField.__init__(self, 'CompleteName', text)
        self.furi = 'contacts:%s' % self.tag

        self.title = self.Title()
        self.given_name = self.GivenName()
        self.first_name = self.FirstName()
        self.middle_name = self.MiddleName()
        self.last_name = self.LastName()
        self.surname = self.Surname()
        self.suffix = self.Suffix()
        self.initials = self.Initials()
        self.nickname = self.Nickname()
        self.full_name = self.FullName()

        self.children = [self.title, self.first_name, self.middle_name,
                         self.last_name, self.suffix, self.initials,
                         self.nickname]

    def write_to_xml(self):
        """
        Return an XML representation of this field.
        """

        children = self.get_children()

        if ((self.value is not None) or (len(self.attrib) > 0) or
                (len(children) > 0)):

            text = self.value_as_xml()
            ats = self.atts_as_xml()
            cs = self.children_as_xml()

            ret = '<t:%s %s>%s%s</t:%s>' % (self.tag, ats, text, cs, self.tag)
            return ret
        else:
            return ''


class SpouseName(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'SpouseName', text)


class Profession(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Profession', text)


class JobTitle(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'JobTitle', text)


class CompanyName(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'CompanyName', text)


class Department(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Department', text)


class Manager(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Manager', text)


class AssistantName(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'AssistantName', text)

# FIXME: The values for these are strings, but might be better represented as
# DateTime objects... Hm


class Birthday(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Birthday', text)


class WeddingAnniversary(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'WeddingAnniversary', text)


class Notes(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'Body', text)
        self.attrib = {
            'BodyType': 'Text',
            # 'IsTruncated' : 'False'
        }
        self.furi = FieldURI(text='item:Body')


class PostalAddress(CField):
    # <t:Street/>
    # <t:City/>
    # <t:State/>
    # <t:CountryOrRegion/>
    # <t:PostalCode/
    class Street(CField):
        def __init__(self, text=None):
            CField.__init__(self, 'Street', text)

    class City(CField):
        def __init__(self, text=None):
            CField.__init__(self, 'City', text)

    class State(CField):
        def __init__(self, text=None):
            CField.__init__(self, 'State', text)

    class CountryOrRegion(CField):
        def __init__(self, text=None):
            CField.__init__(self, 'CountryOrRegion', text)

    class PostalCode(CField):
        def __init__(self, text=None):
            CField.__init__(self, 'PostalCode', text)

    def __init__(self, addr_dict=None):
        CField.__init__(self, 'Entry', addr_dict)
        self.attrib = {
            'Key': None,
        }
        self.street = self.Street()
        self.city = self.City()
        self.state = self.State()
        self.country_region = self.CountryOrRegion()
        self.postal_code = self.PostalCode()
        self.children = [self.street, self.city, self.state,
                         self.country_region, self.postal_code]

        self.tag_field_mapping = {
            'Street': 'street',
            'City': 'city',
            'State': 'state',
            'CountryOrRegion': 'country_region',
            'PostalCode': 'postal_code'
        }

    def key(self):
        return self.attrib['Key']

    # def write_to_xml_update(self):
    #     ret = []
    #     for field in self.children:
    #         ret.append(field.write_to_xml())
    #     return '\n'.join(ret)

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            getattr(self, self.tag_field_mapping[tag]).value = child.text

    def write_to_xml_update(self, field):
        field_class_name = field.__class__.__name__
        s = '<t:SetItemField>'
        s += '\n<t:IndexedFieldURI FieldURI="' \
             'contacts:PhysicalAddress:%s"' % field_class_name
        s += ' FieldIndex="%s"/>' % self.attrib['Key']
        s += '\n<t:Contact>'
        s += '\n  <t:PhysicalAddresses>'
        s += '\n    <t:Entry Key="%s">%s</t:Entry>' % (
            self.attrib['Key'],
            field.write_to_xml_update2())
        s += '\n  </t:PhysicalAddresses>'
        s += '\n</t:Contact>'
        s += '</t:SetItemField>'
        return s

    def write_to_xml_delete(self, field):
        field_class_name = field.__class__.__name__
        s = '<t:DeleteItemField>'
        s += '\n<t:IndexedFieldURI FieldURI="' \
             'contacts:PhysicalAddress:%s"' % field_class_name
        s += ' FieldIndex="%s"/>' % self.attrib['Key']
        s += '</t:DeleteItemField>'
        return s

    def update_or_delete(self):
        ret = []
        for field in self.children:
            if field.value is None:
                # delete
                ret.append(self.write_to_xml_delete(field))
            else:
                # update
                ret.append(self.write_to_xml_update(field))
        return ''.join(ret)


class PostalAddresses(CField):

    def __init__(self, node=None):
        CField.__init__(self, 'PhysicalAddresses')
        self.children = self.entries = []

    def add(self, addr_obj):
        self.entries.append(addr_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, node):
        for child in node:
            addr = PostalAddress()
            for k, v in child.attrib.iteritems():
                addr.add_attrib(k, v)
            addr.populate_from_node(child)
            self.entries.append(addr)

    def get_address_from_key(self, key):
        for addr in self.entries:
            if addr.attrib['Key'] == key:
                return addr

    def has_updates(self):
        return len(self.entries) > 0

    def edit(self, key, update_dict):
        addr = self.get_address_from_key(key)
        if addr:
            for k, v in update_dict.iteritems():
                getattr(addr, k).value = v

    def write_to_xml_update(self):
        """
        See following link for explaination:

        http://stackoverflow.com/questions/13455697/
        updating-a-contacts-physical-address-with-php-ews
        """
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def write_to_xml_delete(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])


class EmailAddresses(CField):

    class Email(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Entry', text)
            self.attrib = {
                'Key': None,
                'Name': None,
                'RoutingType': None,
                'MailBoxType': None
            }

        def key(self):
            return self.attrib['Key']

        def __str__(self):
            return 'Key: %8s  Address: %s' % (self.attrib['Key'], self.value)

        def write_to_xml_update(self):
            s = '<t:SetItemField>'
            s += '\n<t:IndexedFieldURI FieldURI="contacts:EmailAddress"'
            s += ' FieldIndex="%s"/>' % self.attrib['Key']
            s += '\n<t:Contact>'
            s += '\n  <t:EmailAddresses>'
            s += '\n    <t:Entry Key="%s">%s</t:Entry>' % (self.attrib['Key'],
                                                           escape(self.value))
            s += '\n  </t:EmailAddresses>'
            s += '\n</t:Contact>'
            s += '</t:SetItemField>'
            return s

        def write_to_xml_delete(self):
            field_class_name = self.__class__.__name__
            s = '<t:DeleteItemField>'
            s += '\n<t:IndexedFieldURI FieldURI="' \
                 'contacts:EmailAddress:%s"' % field_class_name
            s += ' FieldIndex="%s"/>' % self.attrib['Key']
            s += '</t:DeleteItemField>'
            return s

        def update_or_delete(self):
            ret = []
            if self.value is None:
                # delete
                ret.append(self.write_to_xml_delete())
            else:
                # update
                ret.append(self.write_to_xml_update())
            return ''.join(ret)

    def __init__(self, node=None):
        CField.__init__(self, 'EmailAddresses')
        self.children = self.entries = []
        if node is not None:
            self.populate_from_node(node)

    def populate_from_node(self, node):
        for child in node:
            email = self.Email()
            for k, v in child.attrib.iteritems():
                email.add_attrib(k, v)

            email.value = child.text
            self.entries.append(email)

    def add(self, key, addr):
        if addr is not None:
            email = self.Email()
            email.add_attrib('Key', key)
            email.value = addr
            self.entries.append(email)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def write_to_xml_delete(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def __str__(self):
        s = '%s Addresses: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s


class ImAddresses(CField):

    class Im(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Entry', text)
            self.attrib = {
                'Key': None,
            }

        def key(self):
            return self.attrib['Key']

        def write_to_xml_update(self):
            s = '<t:SetItemField>'
            s += '\n<t:IndexedFieldURI FieldURI="contacts:ImAddress"'
            s += ' FieldIndex="%s"/>' % self.attrib['Key']
            s += '\n<t:Contact>'
            s += '\n  <t:ImAddresses>'
            s += '\n    <t:Entry Key="%s">%s</t:Entry>' % (self.attrib['Key'],
                                                           escape(self.value))
            s += '\n  </t:ImAddresses>'
            s += '\n</t:Contact>'
            s += '</t:SetItemField>'
            return s

        def write_to_xml_delete(self):
            field_class_name = self.__class__.__name__
            s = '<t:DeleteItemField>'
            s += '\n<t:IndexedFieldURI FieldURI="' \
                 'contacts:ImAddress:%s"' % field_class_name
            s += ' FieldIndex="%s"/>' % self.attrib['Key']
            s += '</t:DeleteItemField>'
            return s

        def update_or_delete(self):
            ret = []
            if self.value is None:
                # delete
                ret.append(self.write_to_xml_delete())
            else:
                # update
                ret.append(self.write_to_xml_update())
            return ''.join(ret)

        def __str__(self):
            return 'Key: %8s  Address: %s' % (self.attrib['Key'], self.value)

    def __init__(self, node=None):
        CField.__init__(self, 'ImAddresses')
        self.children = self.entries = []
        if node is not None:
            self.populate_from_node(node)

    def populate_from_node(self, node):
        for child in node:
            im = self.Im()
            for k, v in child.attrib.iteritems():
                im.add_attrib(k, v)

            im.value = child.text
            self.entries.append(im)

    def add(self, key, addr):
        if addr is not None:
            im = self.Im()
            im.add_attrib('Key', key)
            im.value = addr
            self.entries.append(im)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def write_to_xml_delete(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def __str__(self):
        s = '%s Addresses: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s


class PhoneNumbers(CField):

    class Phone(CField):

        def __init__(self, text=None):
            CField.__init__(self, 'Entry', text)
            self.attrib = {
                'Key': None,               # ews.data.PhoneKey
            }

        def key(self):
            return self.attrib['Key']

        def write_to_xml_update(self):
            s = '<t:SetItemField>'
            s += '<t:IndexedFieldURI FieldURI="contacts:PhoneNumber"'
            s += ' FieldIndex="%s" />' % (self.attrib['Key'])
            s += '\n<t:Contact>'
            s += '\n  <t:PhoneNumbers>'
            s += '\n    <t:Entry Key="%s">%s</t:Entry>' % (self.attrib['Key'],
                                                           escape(self.value))
            s += '\n  </t:PhoneNumbers>'
            s += '\n</t:Contact>'
            s += '</t:SetItemField>'
            return s

        def write_to_xml_delete(self):
            field_class_name = self.__class__.__name__
            s = '<t:DeleteItemField>'
            s += '\n<t:IndexedFieldURI FieldURI="' \
                 'contacts:PhoneNumber:%s"' % field_class_name
            s += ' FieldIndex="%s"/>' % self.attrib['Key']
            s += '</t:DeleteItemField>'
            return s

        def update_or_delete(self):
            ret = []
            if self.value is None:
                # delete
                ret.append(self.write_to_xml_delete())
            else:
                # update
                ret.append(self.write_to_xml_update())
            return ''.join(ret)

        def __str__(self):
            return 'Key: %8s  Number: %s' % (self.attrib['Key'], self.value)

    def __init__(self, node=None):
        CField.__init__(self, 'PhoneNumbers')
        self.children = self.entries = []

        if node is not None:
            self.populate_from_node(node)

    def populate_from_node(self, node):
        for child in node:
            phone = self.Phone()
            for k, v in child.attrib.iteritems():
                phone.add_attrib(k, v)

            phone.value = child.text
            self.entries.append(phone)

    def add(self, key, num):
        if num is not None:
            phone = self.Phone()
            phone.add_attrib('Key', key)
            phone.value = num
            self.entries.append(phone)

    def has_updates(self):
        return len(self.entries) > 0

    def write_to_xml_update(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def write_to_xml_delete(self):
        return ''.join([entry.update_or_delete() for entry in self.entries])

    def __str__(self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s


class PostalAddressIndex(CField):
    def __init__(self, text=None):
        CField.__init__(self, 'PostalAddressIndex', text)


class BusinessHomePage(CField):

    def __init__(self, text=None):
        CField.__init__(self, 'BusinessHomePage', text)

##
# Extended Properties
##


class ContactExtendedProperty(ExtendedProperty):
    def __init__(self, node=None, dis_psetid=None, psetid=None,
                 ptag=None, pname=None, pid=None, ptype=None):
        ExtendedProperty.__init__(self, node=node, dis_psetid=dis_psetid,
                                  psetid=psetid, ptag=ptag, pname=pname,
                                  pid=pid, ptype=ptype)

    def write_to_xml_update(self):
        s = '<t:SetItemField>'
        ef = self.efuri.write_to_xml()
        s += ef
        s += '\n<t:Contact>'
        s += '\n  <t:ExtendedProperty>'
        s += '\n      %s' % ef
        s += '\n      <t:Value>%s</t:Value>' % escape(str(self.value))
        s += '\n  </t:ExtendedProperty>'
        s += '\n</t:Contact>'
        s += '</t:SetItemField>'

        return s


class PersonalHomePage(ContactExtendedProperty):

    def __init__(self, node=None, text=None):
        pid = mapitags.PROP_ID(mapitags.PR_PERSONAL_HOME_PAGE)
        ptype = mapitags.PROP_TYPE(mapitags.PR_PERSONAL_HOME_PAGE)
        ContactExtendedProperty.__init__(self, node=node, ptag=pid,
                                         ptype=MapiPropertyTypeType[ptype])
        self.val.value = text

    def write_to_xml(self):
        if self.val.value is not None:
            return ContactExtendedProperty.write_to_xml(self)
        else:
            return ''

    def __str__(self):
        return self.val.value


class Gender(ContactExtendedProperty):

    def __init__(self, node=None, text=GenderType.Unspecified):
        ptag = mapitags.PROP_ID(mapitags.PR_GENDER)
        ptype = mapitags.PROP_TYPE(mapitags.PR_GENDER)
        ContactExtendedProperty.__init__(self, node=node, ptag=ptag,
                                         ptype=MapiPropertyTypeType[ptype])
        self.val.value = str(text)

    def write_to_xml(self):
        if self.val.value is not None:
            return ContactExtendedProperty.write_to_xml(self)
        else:
            return ''

    def __str__(self):
        v = self.val.value
        if v is None or v == 'None':
            return 'Unspecified'
        elif int(v) == GenderType.Female:
            return 'Female'
        elif int(v) == GenderType.Male:
            return 'Male'
        else:
            return 'Unspecified'


class Contact(Item):

    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    def __init__(self, service, parent_fid=None,
                 resp_node=None, mapped_data=None):
        Item.__init__(self, service, parent_fid, resp_node, tag='Contact')

        # self.categories = Categories()
        self.file_as = FileAs()
        self.file_as_mapping = FileAsMapping()
        self.alias = Alias()
        self.complete_name = CompleteName()
        self.display_name = DisplayName()
        self.spouse_name = SpouseName()

        self.job_title = JobTitle()
        self.profession = Profession()
        self.company_name = CompanyName()
        self.department = Department()
        self.manager = Manager()
        self.assistant_name = AssistantName()

        self.birthday = Birthday()
        self.anniversary = WeddingAnniversary()

        self.notes = Notes()
        self.emails = EmailAddresses()
        self.ims = ImAddresses()
        self.physical_addresses = PostalAddresses()
        self.phones = PhoneNumbers()
        self.business_home_page = BusinessHomePage()
        self.postal_address_index = PostalAddressIndex()

        self.gender = Gender()
        self.personal_home_page = PersonalHomePage()

        if mapped_data is not None:
            self._init_from_mapped_data(mapped_data)
        elif resp_node is not None:
            self._init_from_resp()

    def _init_from_mapped_data(self, data):
        "Create a Contact record from the mapped data"
        if not data:
            return

        # @TODO factorize with init_from_resp maybe ?
        if 'exchange_id' not in data:
            self.categories.add('odoo')
        for key in data.keys():
            if key == 'exchange_id' and data.get('exchange_id'):
                self.itemid.set(data.get('exchange_id', ''))

            if key == 'change_key' and data.get('change_key'):
                self.change_key.set(data.get('change_key'))

            if key == 'FileAs' and data.get('FileAs'):
                self.file_as.set(data.get('FileAs'))

            if key == 'Alias' and data.get('Alias'):
                self.file_as.set(data.get('Alias'))

            if key == 'Name' and data.get('Name'):
                self.display_name.set(data.get('Name'))

            if key == 'title' and data.get('title'):
                self.complete_name.title.set(data.get('title'))

            if key == 'DisplayName' and data.get('DisplayName'):
                self.complete_name.first_name.set(data.get('DisplayName'))

            if key == 'FirstName' and data.get('FirstName'):
                self.complete_name.first_name.set(data.get('FirstName'))
                self.complete_name.given_name.set(data.get('FirstName'))

            if key == 'LastName' and data.get('LastName'):
                self.complete_name.surname.set(data.get('LastName'))
                self.complete_name.last_name.set(data.get('LastName'))

            if key == 'JobTitle' and data.get('JobTitle'):
                self.job_title.set(data.get('JobTitle'))

            if key == 'CompanyName' and data.get('CompanyName'):
                self.company_name.set(data.get('CompanyName'))

            if key == 'Phone' and data.get('Phone'):
                self.phones.add('BusinessPhone', data.get('Phone'))

            if key == 'Fax' and data.get('Fax'):
                self.phones.add('BusinessFax', data.get('Fax'))

            if key == 'Mobile' and data.get('Mobile'):
                self.phones.add('MobilePhone', data.get('Mobile'))

            if key == 'Email' and data.get('Email'):
                self.emails.add('EmailAddress1', data.get('Email'))

            if key == 'BusinessHomePage' and data.get('BusinessHomePage'):
                self.business_home_page.text = data.get('BusinessHomePage')

            self._firstname = self.complete_name.first_name.value or ''
            self._lastname = self.complete_name.last_name.value or ''
            self._displayname = self.display_name.value or ''

    def _init_from_resp(self):
        if self.resp_node is None:
            return

        rnode = self.resp_node
        for child in rnode:
            tag = unQName(child.tag)

            if tag == 'FileAs':
                self.file_as.value = child.text
            elif tag == 'Alias':
                self.alias.value = child.text
            elif tag == 'SpouseName':
                self.spouse_nam.value = child.text
            elif tag == 'JobTitle':
                self.job_title.value = child.text
            elif tag == 'CompanyName':
                self.company_name.value = child.text
            elif tag == 'Department':
                self.department.value = child.text
            elif tag == 'Manager':
                self.manager.value = child.text
            elif tag == 'AssistantName':
                self.assistant_name.value = child.text
            elif tag == 'Birthday':
                self.birthday.value = child.text
            elif tag == 'WeddingAnniversary':
                self.anniversary.value = child.text
            elif tag == 'GivenName':
                self.complete_name.given_name.value = child.text
            elif tag == 'Surname':
                self.complete_name.surname.value = child.text
            elif tag == 'Initials':
                self.complete_name.initials.value = child.text
            elif tag == 'DisplayName':
                self.display_name.value = child.text
            elif tag == 'Body':
                # FIXME: We are assuming a text body type, but they could
                # contain html or other types as well... Oh, well.
                self.notes.value = child.text
            elif tag == 'EmailAddresses':
                self.emails.populate_from_node(child)
            elif tag == 'ImAddresses':
                self.ims.populate_from_node(child)
            elif tag == 'PhysicalAddresses':
                self.physical_addresses.populate_from_node(child)
            elif tag == 'PhoneNumbers':
                self.phones.populate_from_node(child)
            elif tag == 'BusinessHomePage':
                self.business_home_page.value = child.text
            elif tag == 'ExtendedProperty':
                self.add_extended_property(node=child)

        n = rnode.find('CompleteName')
        if n is not None:
            rnode = n

        fts = self._find_text_safely
        self.complete_name.title.value = fts(rnode, 'Title')
        self.complete_name.first_name.value = fts(rnode, 'FirstName')
        self.complete_name.middle_name.value = fts(rnode, 'MiddleName')
        self.complete_name.last_name.value = fts(rnode, 'LastName')
        self.complete_name.suffix.value = fts(rnode, 'Suffix')
        self.complete_name.initials.value = fts(rnode, 'Initials')
        self.complete_name.full_name.value = fts(rnode, 'FullName')
        self.complete_name.nickname.value = fts(rnode, 'Nickname')

        # It's a bit hard to understand why the hell they have so many
        # variants for the same stupid information... Oh well, let's just
        # have a few handy shortcuts for the information that matters
        if self.complete_name.first_name.value:
            self._firstname = self.complete_name.first_name.value
        else:
            self._firstname = self.complete_name.given_name.value

        if self.complete_name.last_name.value:
            self._lastname = self.complete_name.last_name.value
        else:
            self._lastname = self.complete_name.surname.value

        if self._firstname is None:
            self._firstname = ''

        if self._lastname is None:
            self._lastname = ''

        if self.display_name.value:
            self._displayname = self.display_name.value
        elif self.complete_name.full_name.value:
            self._displayname = self.complete_name.full_name.value
        else:
            self._displayname = self._firstname + ' ' + self._lastname

        # yet to support following which are really multi-valued properties
        # - Companies
        # - PhysicalAddresses
        # - ImAddresses

        # yet to support following which are not returned normally and have
        # to be dealt with as extended properties.
        # - gender
        # - LastModifiedTime

    ##
    # Inherited methods. For doc see item.py
    ##

    def add_tagged_property(self, node=None, tag=None, value=None):
        if node is not None and tag is None:
            uri = node.find(QName_T('ExtendedFieldURI'))
            tag = ExtendedProperty.get_prop_tag_from_xml(uri)
            value = node.find(QName_T('Value')).text

        eprop = None

        if tag == mapitags.PR_LAST_MODIFICATION_TIME:
            self.last_modified_time = LastModifiedTime(node=node, text=value)
            eprop = self.last_modified_time
        elif tag == mapitags.PR_GENDER:
            self.gender = Gender(node=node, text=value)
            eprop = self.gender
        elif tag == mapitags.PR_PERSONAL_HOME_PAGE:
            self.personal_home_page = PersonalHomePage(node=node, text=value)
            eprop = self.personal_home_page
        else:
            eprop = ExtendedProperty(node=node, ptag=tag)
            eprop.value = value
            self.eprops.append(eprop)
            self.eprops_tagged[tag] = eprop

    # Override Field.get_chidren to refresh the children array every time
    def get_children(self):
        cn = self.complete_name
        # Note that children is used for generating xml representation of
        # this contact for CreateItem and update operations. The order of
        # these fields is critical. I know, it's crazy.
        self.children = [self.notes] + self.eprops + [
            self.categories,
            self.gender,
            self.personal_home_page, self.file_as, self.file_as_mapping,
            self.display_name, cn.given_name, cn.initials,
            cn.middle_name, cn.nickname, self.company_name,
            self.emails, self.physical_addresses, self.phones,
            self.assistant_name,
            self.birthday, self.business_home_page,
            self.department, self.ims, self.job_title,
            self.manager, self.postal_address_index, self.profession,
            self.spouse_name, cn.surname, self.anniversary, self.alias]

        return self.children

    def save(self):
        if self.itemid.value is None:
            self.service.CreateItems(self.parent_fid.value, [self])
        else:
            self.service.UpdateItems([self])

    def __str__(self):
        lmt_tag = mapitags.PR_LAST_MODIFICATION_TIME

        s = 'ItemId: %s' % self.itemid
        s += '\nChangeKey: %s' % self.change_key
        s += '\nCreated: %s' % self.created_time
        if lmt_tag in self.eprops_tagged:
            s += '\nLast Modified: %s' % self.eprops_tagged[lmt_tag].val.value
        s += '\nName: %s' % self._displayname
        s += '\nGEnder: %s' % self.gender
        s += '\nPhones: %s' % self.phones
        s += '\nEmails: %s' % self.emails
        s += '\nIms: %s' % self.ims
        s += '\nNotes: %s' % self.notes
        s += '\nJob title: %s' % self.job_title.value

        return s

# The XML Schema for a EWS Contact, taken from:
# http://msdn.microsoft.com/en-us/library/office/aa581315(v=exchg.150).aspx
##
# <Contact>
# <MimeContent/>
# <ItemId/>
# <ParentFolderId/>
# <ItemClass/>
# <Subject/>
# <Sensitivity/>
# <Body/>
# <Attachments/>
# <DateTimeReceived/>
# <Size/>
# <Categories/>
# <Importance/>
# <InReplyTo/>
# <IsSubmitted/>
# <IsDraft/>
# <IsFromMe/>
# <IsResend/>
# <IsUnmodified/>
# <InternetMessageHeaders/>
# <DateTimeSent/>
# <DateTimeCreated/>
# <ResponseObjects/>
# <ReminderDueBy/>
# <ReminderIsSet/>
# <ReminderMinutesBeforeStart/>
# <DisplayCc/>
# <DisplayTo/>
# <HasAttachments/>
# <ExtendedProperty/>
# <Culture/>
# <EffectiveRights/>
# <LastModifiedName/>
# <LastModifiedTime/>
# <IsAssociated/>
# <WebClientReadFormQueryString/>
# <WebClientEditFormQueryString/>
# <ConversationId/>
# <UniqueBody/>
# <FileAs/>
# <FileAsMapping/>
# <DisplayName/>
# <GivenName/>
# <Initials/>
# <MiddleName/>
# <Nickname/>
# <CompleteName>
#     <Title/>
#     <FirstName/>
#     <MiddleName/>
#     <LastName/>
#     <Suffix/>
#     <Initials/>
#     <FullName/>
#     <Nickname/>
#     <YomiFirstName/>
#     <YomiLastName/>
# </CompleteName>
# <CompanyName/>
# <EmailAddresses/>
# <PhysicalAddresses/>
# <PhoneNumbers/>
# <AssistantName/>
# <Birthday/>
# <BusinessHomePage/>
# <Children/>
# <Companies/>
# <ContactSource/>
# <Department/>
# <Generation/>
# <ImAddresses/>
# <JobTitle/>
# <Manager/>
# <Mileage/>
# <OfficeLocation/>
# <PostalAddressIndex/>
# <Profession/>
# <SpouseName/>
# <Surname/>
# <WeddingAnniversary/>
# <HasPicture/>
# <PhoneticFullName/>
# <PhoneticFirstName/>
# <PhoneticLastName/>
# <Alias/>
# <Notes/>
# <Photo/>
# <UserSMIMECertificate/>
# <MSExchangeCertificate/>
# <DirectoryId/>
# <ManagerMailbox/>
# <DirectReports/>
# </Contact>
