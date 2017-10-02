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

from abc import ABCMeta, abstractmethod
from pyews.soap import SoapClient, QName_M, QName_T, unQName
import pyews.soap as soap
import pyews.utils as utils
from pyews.ews import mapitags
from pyews.ews.data import MapiPropertyTypeType, MapiPropertyTypeTypeInv
from pyews.ews.data import (SensitivityType,
                            ImportanceType,
                            )
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
import logging
import pdb

gnd = SoapClient.get_node_detail
_logger = logging.getLogger(__name__)
EXCHANGE_DATETIME_FORMAT = '%Y-%m-%dT%H:%M:%SZ'


class ReadOnly:

    """
    When applied as a Mixin, this class will ensure that no XML is generated
    when this field is written out as part of the containing Item. Usage for
    this:

    class LastModifiedTime(ReadOnly, ExtendedProperty):
        pass

    Note that for this to work in python (given the rules of method resolution
    order for multiple inheritance, the ReadOnly Mixin should come to the left
    of Field class or any its derived classes.
    """
    __metaclass__ = ABCMeta

    def write_to_xml(self):
        return ''

    def write_to_xml_update(self):
        return ''


class Field:

    """
    Represents an XML Element
    """
    __metaclass__ = ABCMeta

    def __init__(self, tag=None, text=None, boolean=False):
        self.tag = tag
        self.value = text
        self.attrib = {}
        self.children = []
        self.read_only = False
        self.boolean = boolean

        # furi is used when this field needs to be used as part of an update
        # item method
        self.furi = ('item:%s' % tag) if tag else None

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, val):
        self._value = val

    def add_attrib(self, key, val):
        self.attrib.update({key: val})

    def atts_as_xml(self):
        ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
        return ' '.join(ats)

    def value_as_xml(self):
        if self.value is not None:
            if self.boolean:
                value = escape(unicode(self.value).lower())
            else:
                value = escape(unicode(self.value))
            return value
        return ''

    def children_as_xml(self):
        self.children = self.get_children()
        # xmls = [x.write_to_xml() for x in self.children]
        xmls = [x.write_to_xml() for x in self.children if not (isinstance(
            x, basestring) or x is None)]

        return '\n'.join([y for y in xmls if y is not None])

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

    def write_to_xml_get(self):
        """
        Presently only makes sense for certain ExtendedProperties
        """

        ats = self.atts_as_xml()
        ret = '<t:%s %s/>' % (self.tag, ats)
        return ret

    def write_to_xml_update(self):
        if (self.value is not None):

            text = self.value_as_xml()
            ats = self.atts_as_xml()

            ret = '<t:SetItemField>'
            ret += '<t:FieldURI FieldURI="%s"/>' % self.furi
            ret += '<t:Item>'
            ret += '<t:%s %s>%s</t:%s>' % (self.tag, ats, text, self.tag)
            ret += '</t:Item>'
            ret += '</t:SetItemField>'

            return ret
        else:
            return ''

    def has_updates(self):
        return not (self.value is None or
                    (isinstance(self.value, list) and len(self.value) == 0))

    def get_children(self):
        return self.children

    def set(self, value):
        self.value = value

    def __str__(self):
        return self.value if self.value is not None else ""


class FieldURI(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'FieldURI', text)


class ItemId(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'ItemId', text)


class ChangeKey(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'ChangeKey', text)


class ParentFolderId(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'ParentFolderId', text)
        self.furi = 'folder:ParentFolderId'


class ParentFolderChangeKey(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'ParentFolderChangeKey', text)


class ItemClass(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'ItemClass', text)
        self.furi = 'item:ItemClass'


class DateTimeCreated(Field):

    def __init__(self, text=None):
        Field.__init__(self, 'DateTimeCreated', text)


class PropVariant:
    UNKNOWN = 0
    TAGGED = 1
    NAMED_INT = 2
    NAMED_STR = 3


class Content(Field):
    def __init__(self, text=None):
        Field.__init__(self, 'Content', text)


class FileAttachment(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa580492(v=exchg.140).aspx
    """

    class AttachmentId(Field):
        """
    https://msdn.microsoft.com/en-us/library/office/aa580987(v=exchg.140).aspx
        """
        def __init__(self, text=None):
            Field.__init__(self, 'AttachmentId', text)

    class Name(Field):
        def __init__(self, text=None):
            Field.__init__(self, 'Name', text)

    class ContentType(Field):
        def __init__(self, text=None):
            Field.__init__(self, 'ContentType', text)

    def __init__(self, service, node=None):
        Field.__init__(self, 'FileAttachment')
        self.service = service

        self.attachment_id = self.AttachmentId(self.service)
        self.name = self.Name()
        self.content_type = self.ContentType()
        self.content = Content()

        self.tag_field_mapping = {
            'AttachmentId': 'attachment_id',
            'Name': 'name',
            'ContentType': 'content_type',
            'Content': 'content',
        }

        self.children = [self.name, self.content]

    def populate_from_node(self, node):
        for child in node:
            tag = unQName(child.tag)
            if tag == self.attachment_id.tag:
                self.attachment_id.set(child.attrib['Id'])
            elif tag == self.content.tag:
                self.content.set(child.text)
            elif tag == self.name.tag:
                self.name.set(child.text)
            else:
                continue
                # getattr(self, self.tag_field_mapping[tag]).value = child.text


class Attachments(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa564869(v=exchg.140).aspx

<t:Attachments>
    <t:FileAttachment>
        <t:AttachmentId Id="AAAQAGMxb2Rvb0BoaWdoY2..."/>
        <t:Name>CalendarItem.xml</t:Name>
        <t:ContentType>text/xml</t:ContentType>
    </t:FileAttachment>
</t:Attachments>
    """
    def __init__(self, service, node=None):
        Field.__init__(self, 'Attachments')
        self.service = service
        self.entries = []
        self.children = self.get_children()
        if node is not None:
            self.populate_from_node(service, node)

    def add(self, att_obj):
        self.entries.append(att_obj)

    def get_children(self):
        return self.entries

    def populate_from_node(self, service, node):
        for child in node:
            tag = unQName(child.tag)
            if tag != 'FileAttachment':
                continue
            att = FileAttachment(service)
            att.populate_from_node(child)
            if att.content.value is None:
                # retrieve attachment and set content with it
                cnt = service.GetAttachments(
                    [att.attachment_id.value])
                att.content.set(cnt[0].value)
            self.entries.append(att)

    def write_to_xml(self):
        return ''


class ExtendedProperty(Field):

    class ExtendedFieldURI(Field):

        def __init__(self, node=None, dis_psetid=None, psetid=None,
                     ptag=None, pname=None, pid=None, ptype=None):
            Field.__init__(self, 'ExtendedFieldURI')

            # See here for explanation of each of these and what the
            # constraints are for mixing and matchig these fields.
            # http://msdn.microsoft.com/en-us/library/office/aa564843.aspx

            # Note - All of these values will be strings (if not None)
            self.attrib = {
                'DistinguishedPropertySetId': dis_psetid,
                'PropertySetId': psetid,
                'PropertyTag': ptag,
                'PropertyName': pname,
                'PropertyId': pid,
                'PropertyType': ptype,
            }

            if node is not None:
                self.init_from_xml(node)

        def init_from_xml(self, node):
            """
            Node is a parsed xml node (Element). Extract the data that we can
            from the node.
            """

            uri = node.find(QName_T('ExtendedFieldURI'))
            if uri is None:
                logging.debug(
                    'ExtendedProperty.init_from_xml(): no child node ' +
                    'ExtendedFieldURI in node: %s',
                    pretty_xml(node))
            else:
                self.attrib.update(uri.attrib)

            # The PropertyId attribute needs to be an integer for easier
            # processing, so fix it
            pid = self.attrib['PropertyId']
            if pid is not None:
                self.attrib['PropertyId'] = utils.safe_int(pid)

    class Value(Field):

        def __init__(self, text=None):
            Field.__init__(self, 'Value')

    def __init__(self, node=None, dis_psetid=None, psetid=None,
                 ptag=None, pname=None, pid=None, ptype=None):
        """
        If node is not None, it should be a parsed XML element pointing to an
        Extended Property element
        """
        # This guy needs to be here as we have overriden the .value()
        self.val = self.Value()

        Field.__init__(self, 'ExtendedProperty')
        self.efuri = self.ExtendedFieldURI(node, dis_psetid, psetid,
                                           ptag, pname, pid, ptype)

        # FIXME: We can have a multi-valued property as well.
        if node is not None:
            self.value = node.find(QName_T('Value')).text

        self.children = [self.efuri, self.val]

    ##
    # Getter and Setters.
    ##

    @property
    def value(self):
        return self.val.value

    @value.setter
    def value(self, text):
        self.val.value = text

    @property
    def psetid(self):
        return self.efuri.attrib['PropertySetId']

    @psetid.setter
    def psetid(self, val):
        self.efuri.attrib['PropertySetId'] = val

    @property
    def pid(self):
        return self.efuri.attrib['PropertyId']

    @pid.setter
    def pid(self, val):
        self.efuri.attrib['PropertyId'] = val

    @property
    def pname(self):
        return self.efuri.attrib['PropertyName']

    @pname.setter
    def pname(self, val):
        self.efuri.attrib['PropertyName'] = val

    ##
    # Overriding inherted methods
    ##

    def value_as_xml(self):
        # Any attribute that has a 'Value' child should have its own value
        # printed out as text in the XML represented. Note for extended
        # properties self.value returns self.val.value for ease of access. So
        # if we do not do this the value will go out twide

        return ''

    def write_to_xml_get(self):
        """
        Presently only makes sense for certain ExtendedProperties
        """

        return self.efuri.write_to_xml_get()

    def write_to_xml_update(self):
        s = ''

        ef = self.efuri.write_to_xml()
        s += ef
        s += '\n<t:Item>'
        s += '\n  <t:ExtendedProperty>'
        s += '\n      %s' % ef
        s += '\n      <t:Value>%s</t:Value>' % escape(str(self.value))
        s += '\n  </t:ExtendedProperty>'
        s += '\n</t:Item>'

        return s

    def get_variant(self):
        """
        Outlook has three types of properties - Tagged Properties, Named
        Properites with Numeric Identifiers, and Named Properties with
        string identifiers. This method will identify which variant this
        property is. It will return one of the values defined in the
        PropVariant class above.
        """

        all_none = True
        for v in self.efuri.attrib.values():
            if v is not None:
                all_none = False
                break
        if all_none:
            return PropVariant.UNKNOWN

        if (self.efuri.attrib['PropertyTag'] is not None and
                self.efuri.attrib['PropertyType'] is not None and
                self.efuri.attrib['PropertySetId'] is None and
                self.efuri.attrib['DistinguishedPropertySetId'] is None and
                self.efuri.attrib['PropertyName'] is None and
                self.efuri.attrib['PropertyId'] is None):
            return PropVariant.TAGGED

        if (self.efuri.attrib['PropertyTag'] is None and
                self.efuri.attrib['PropertyType'] is not None and
                (self.efuri.attrib['PropertySetId'] is not None or
                 self.efuri.attrib['DistinguishedPropertySetId'] is not None
                 ) and
                self.efuri.attrib['PropertyName'] is None and
                self.efuri.attrib['PropertyId'] is not None):
            return PropVariant.NAMED_INT

        if (self.efuri.attrib['PropertyTag'] is None and
                self.efuri.attrib['PropertyType'] is not None and
                (self.efuri.attrib['PropertySetId'] is not None or
                 self.efuri.attrib['DistinguishedPropertySetId'] is not None
                 ) and
                self.efuri.attrib['PropertyName'] is not None and
                self.efuri.attrib['PropertyId'] is None):
            return PropVariant.NAMED_STR

        return PropVariant.UNKNOWN

    def get_prop_tag(self):
        """Return a tag version of the current property. Note that the
        name tag is used in different ways by MS in different places. Here
        we are talking about the combined property_tag and property_type
        entity."""

        pt = self.efuri.attrib['PropertyTag']
        base = 16 if pt[0:2] == "0x" else 10
        pid = int(pt, base)
        ptype = int(MapiPropertyTypeTypeInv[self.efuri.attrib['PropertyType']])

        return mapitags.PROP_TAG(ptype, pid)

    def set(self, value):
        self.value = value

    ##
    # Some helper methods to easily "recognize" the extended properties
    ##

    @staticmethod
    def get_variant_from_xml(xml_node):
        """
        If the element is a tagged property then returns (True, MAPITag)
        otherwise it will be (False, ignore)
        """
        if (len(xml_node.attrib) == 2 and
                ('PropertyTag' in xml_node.attrib and
                 'PropertyType' in xml_node.attrib
                 )):
            return PropVariant.TAGGED

        if (len(xml_node.attrib) == 3 and
                (('PropertySetId' in xml_node.attrib or
                 'DistinguishedPropertySetId' in xml_node.attrib) and
                 'PropertyId' in xml_node.attrib and
                 'PropertyType' in xml_node.attrib)):
            return PropVariant.NAMED_INT

        if (len(xml_node.attrib) == 3 and
                (('PropertySetId' in xml_node.attrib or
                 'DistinguishedPropertySetId' in xml_node.attrib) and
                 'PropertyName' in xml_node.attrib and
                 'PropertyType' in xml_node.attrib)):
            return PropVariant.NAMED_STR

        return PropVariant.UNKNOWN

    @staticmethod
    def get_prop_tag_from_xml(xml_node):
        """
        Given xml_node which corresponds to a parsed XML element for a tagged
        property get the mapitag for this element. It is an error to call this
        method on non-tagged property elements
        """

        tp = (len(xml_node.attrib) == 2 and
              'PropertyTag' in xml_node.attrib and
              'PropertyType' in xml_node.attrib)

        assert tp

        pid = utils.safe_int(xml_node.attrib['PropertyTag'])
        ptype = xml_node.attrib['PropertyType']
        tag = mapitags.PROP_TAG(MapiPropertyTypeTypeInv[ptype], pid)

        return tag

    @staticmethod
    def get_named_int_from_xml(xml_node):
        """
        Given xml_node which corresponds to a parsed XML element for a tagged
        property get the mapitag for this element. It is an error to call this
        method on non-tagged property elements
        """

        tp = (len(xml_node.attrib) == 3 and
              (('PropertySetId' in xml_node.attrib or
                'DistinguishedPropertySetId' in xml_node.attrib
                ) and
               'PropertyId' in xml_node.attrib and
               'PropertyType' in xml_node.attrib)
              )

        assert tp

        if 'PropertySetId' in xml_node.attrib:
            psetid = xml_node.attrib['PropertySetId']
        else:
            psetid = None

        if 'DistinguishedPropertySetId' in xml_node.attrib:
            dpsetid = xml_node.attrib['DistinguishedPropertySetId']
        else:
            dpsetid = None

        pid = utils.safe_int(xml_node.attrib['PropertyId'])

        return dpsetid, psetid, pid

    @staticmethod
    def get_named_str_from_xml(xml_node):
        """
        Given xml_node which corresponds to a parsed XML element for a tagged
        property get the mapitag for this element. It is an error to call this
        method on non-tagged property elements
        """

        tp = (len(xml_node.attrib) == 3 and
              (('PropertySetId' in xml_node.attrib or
                'DistinguishedPropertySetId' in xml_node.attrib
                ) and
               'PropertyName' in xml_node.attrib and
               'PropertyType' in xml_node.attrib)
              )

        assert tp

        if 'PropertySetId' in xml_node.attrib:
            psetid = xml_node.attrib['PropertySetId']
        else:
            psetid = None

        if 'DistinguishedPropertySetId' in xml_node.attrib:
            dpsetid = xml_node.attrib['DistinguishedPropertySetId']
        else:
            dpsetid = None

        pname = xml_node.attrib['PropertyName']

        return dpsetid, psetid, pname


class LastModifiedTime(ReadOnly, ExtendedProperty):

    def __init__(self, node=None, text=None):
        ptag = mapitags.PROP_ID(mapitags.PR_LAST_MODIFICATION_TIME)
        ptype = mapitags.PROP_TYPE(mapitags.PR_LAST_MODIFICATION_TIME)
        ExtendedProperty.__init__(self, node=node, ptag=ptag, ptype=ptype)


class Categories(Field):
    class Category(Field):
        def __init__(self, text=None):
            Field.__init__(self, 'String', text)

        def __str__(self):
            return 'Category %s' % (self.value)

    def __init__(self, node=None):
        Field.__init__(self, 'Categories')
        self.children = self.entries = []
        if node is not None:
            self.populate_from_node(node)

    def populate_from_node(self, node):
        for child in node:
            self.add(child.text)

    def already_exists(self, categ_str):
        for entry in self.entries:
            if entry.value == categ_str:
                return True
        return False

    def add(self, categ_str):
        if categ_str is not None:
            if not self.already_exists(categ_str):
                categ = self.Category()
                categ.value = categ_str
                self.entries.append(categ)

    def has_updates(self):
        return len(self.entries) > 0


class Subject(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa565100(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        Field.__init__(self, 'Subject', text)


class HasAttachments(Field):
    """
    """
    def __init__(self, text=None):
        Field.__init__(self, 'HasAttachments', text)


class Sensitivity(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa565687(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = SensitivityType._props_values()
        err = 'Sensitivity is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        Field.__init__(self, 'Sensitivity', text)


class Importance(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa563467(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        val_list = ImportanceType._props_values()
        err = 'Importance is not in the list %s' % val_list
        assert (text is None or text in val_list), err
        Field.__init__(self, 'Importance', text)


class Body(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa581015(v=exchg.140).aspx
    """
    def __init__(self, type='HTML', text=None):
        Field.__init__(self, 'Body', text)
        self.text_type = type
        self.attrib['BodyType'] = self.text_type


class ReminderIsSet(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa566410%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        Field.__init__(self, 'ReminderIsSet', text, boolean=True)


class IsDraft(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa566410%28v=exchg.140%29.aspx
    """
    def __init__(self, text=None):
        Field.__init__(self, 'IsDraft', text, boolean=True)


class ReminderDueBy(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa565894(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        Field.__init__(self, 'ReminderDueBy', text)


class ReminderMinutesBeforeStart(Field):
    """
https://msdn.microsoft.com/en-us/library/office/aa581305(v=exchg.140).aspx
    """
    def __init__(self, text=None):
        Field.__init__(self, 'ReminderMinutesBeforeStart', text)


class Item(Field):

    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    __metaclass__ = ABCMeta

    def __init__(self, service, parent_fid=None, resp_node=None, tag='Item',
                 additional_properties=None,
                 additional_boolean_fields_tag=None):
        Field.__init__(self, tag=tag)

        self.service = service  # Exchange service object
        self.resp_node = resp_node

        self.parent_fid = ParentFolderId(parent_fid)
        self.parent_fck = None

        # Now initialize some of the properties to default values
        self.itemid = ItemId()
        self.item_class = ItemClass()
        self.change_key = ChangeKey()
        self.created_time = DateTimeCreated()
        self.last_modified_time = LastModifiedTime()

        self.categories = Categories()
        self.sensitivity = Sensitivity()
        self.importance = Importance()
        self.subject = Subject()
        self.body = Body()
        self.attachments = Attachments(service)
        self.has_attachments = HasAttachments()
        self.reminder_due_by = ReminderDueBy()
        self.is_reminder_set = ReminderIsSet()
        self.is_draft = IsDraft()
        self.reminder_minutes_before_start = ReminderMinutesBeforeStart()

        self.tag_property_map = [
            (self.itemid.tag, self.itemid),
            (self.item_class.tag, self.item_class),
            (self.parent_fid.tag, self.parent_fid),
            (self.created_time.tag, self.created_time),
            (self.last_modified_time.tag, self.last_modified_time),
            (self.categories.tag, self.categories),
            (self.sensitivity.tag, self.sensitivity),
            (self.importance.tag, self.importance),
            (self.is_draft.tag, self.is_draft),
            (self.subject.tag, self.subject),
            (self.body.tag, self.body),
            (self.attachments.tag, self.attachments),
            (self.has_attachments.tag, self.has_attachments),

            (self.reminder_due_by.tag, self.reminder_due_by),
            (self.is_reminder_set.tag, self.is_reminder_set),
            (self.reminder_minutes_before_start.tag,
             self.reminder_minutes_before_start),
        ]
        if additional_properties is not None:
            self.tag_property_map += additional_properties

        self.mapping_dict_tag_obj = dict(self.tag_property_map)

        # Tags starting with 'Is' are considered as Boolean fields
        # for other fields that must be considered as Boolean ones,
        # add them in the list below
        self.boolean_fields_tag = [
            self.is_reminder_set.tag,
            self.has_attachments.tag,
            self.is_draft.tag
        ]
        if additional_boolean_fields_tag is not None:
            self.boolean_fields_tag += additional_boolean_fields_tag

        self.eprops = []
        self.eprops_tagged = {}
        self.eprops_named_str = {}
        self.eprops_named_int = {}

        if self.resp_node is not None:
            self._init_base_fields_from_resp(resp_node)

    ##
    # First the abstract methods that will be implementd by sub classes
    ##

    @abstractmethod
    def add_tagged_property(self, node=None, tag=None, value=None):
        """
        Add a tagged property either from an XML node or from the individual
        components. If node is not None the other parameters are ignored.
        """
        raise NotImplementedError

    ##
    # Next, the non-abstract external methods
    ##

    def _get_atts(self, child):
        return []

    def _get_sets(self, child):
        if child.has_updates():
            return [child]
        else:
            return []

    def _get_dels(self, child):
        if not child.has_updates():
            return [child]
        else:
            return []

    def get_updates(self):
        """
        Returns a list of child elements that need to be included
        in an UpdateItem call. Returns three arrays as a tupple: (adds, sets,
        dels)
        """

        atts = []
        sets = []
        dels = []

        for child in self.get_children():
            atts.extend(self._get_atts(child))
            sets.extend(self._get_sets(child))
            dels.extend(self._get_dels(child))

        # print 'Sets: ', sets
        # print 'Dels: ', dels
        return atts, sets, dels

    def get_extended_properties(self):
        return self.eprops

    def add_extended_property(self, node):
        """
        Parse the XML element pointed to by node, figure out the type of
        property this is, initialize the property of the right type and then
        insert that into the self.eprops member variable
        """

        uri = node.find(QName_T('ExtendedFieldURI'))
        if uri is None:
            logging.error('ExtendedProperty.init_from_xml(): no child node ' +
                          'ExtendedFieldURI in node: %s',
                          pretty_xml(node))
            return

        # Look for known extended properties
        v = ExtendedProperty.get_variant_from_xml(uri)
        if v == PropVariant.TAGGED:
            self.add_tagged_property(node=node)
        elif v == PropVariant.NAMED_INT:
            self.add_named_int_property(node=node)
        elif v == PropVariant.NAMED_STR:
            self.add_named_str_property(node=node)
        else:
            logging.debug('Unrecognized ExtendedProp Variant. Useless')
            self.eprops.append(ExtendedProperty(node=node))

    def add_named_str_property(self, node=None, psetid=None, pname=None,
                               ptype=None, value=None):
        eprop = ExtendedProperty(node=node, psetid=psetid, pname=pname,
                                 ptype=ptype)
        if value is not None:
            eprop.value = value

        self.eprops.append(eprop)
        if eprop.psetid in self.eprops_named_str:
            self.eprops_named_str[eprop.psetid].update({eprop.pname: eprop})
        else:
            self.eprops_named_str[eprop.psetid] = {eprop.pname: eprop}

    def add_named_int_property(self, node=None, psetid=None, pid=None,
                               ptype=None, value=None):
        eprop = ExtendedProperty(node=node, psetid=psetid, pid=pid,
                                 ptype=ptype)

        if value is not None:
            eprop.value = value

        self.eprops.append(eprop)
        if eprop.psetid in self.eprops_named_int:
            self.eprops_named_int[eprop.psetid].update({eprop.pid: eprop})
        else:
            self.eprops_named_int[eprop.psetid] = {eprop.pid: eprop}

    def get_tagged_property(self, tag):
        """
        Return the Tagged ExtendedProperty object that corresponds to the
        specified tag. Returns None if no eprop exists with specified tag
        """

        try:
            return self.eprops_tagged[tag]
        except KeyError, e:
            return None

    def get_named_str_property(self, psetid, pname):
        """
        Return the Named String ExtendedProperty object that corresponds to
        the specified psetid and . Returns None if no such eprop exists
        """

        try:
            return self.eprops_named_str[psetid][pname]
        except KeyError, e:
            return None

    def get_named_int_property(self, psetid, pid):
        """
        Return the Named Integer ExtendedProperty object that corresponds to
        the specified psetid and . Returns None if no such eprop exists
        """

        try:
            return self.eprops_named_int[psetid][pid]
        except KeyError, e:
            logging.debug('Named int prop missing-psetid : %s, pid: 0x%x',
                          psetid, pid)
            return None

    ##
    # Getters and Setters
    ##

    @property
    def parent_folder_id(self):
        return self.parent_folder_id

    @parent_folder_id.setter
    def parent_folder_id(self, val):
        self.parent_folder_id = val

    ##
    # Finally, some internal methods and helper functions
    ##

    def _find_text_safely(self, elem, node):
        r = elem.find(node)
        return r.value if r else None

    def _process_tag(self, child, tag):
        if tag == 'ItemId':
            self.itemid.set(child.attrib['Id'])
            self.change_key.set(child.attrib['ChangeKey'])
        elif tag == 'ParentFolderId':
            self.parent_fid = ParentFolderId(child.attrib['Id'])
            self.parent_fck = ParentFolderChangeKey(
                child.attrib['ChangeKey'])
        elif tag == 'Body':
            # check text type (BodyType attribute on Body tag)
            attribs = child.attrib
            self.mapping_dict_tag_obj[tag].text_type = (
                attribs.get('BodyType', 'HTML')
            )
            setattr(self.mapping_dict_tag_obj[tag],
                    'value',
                    child.text)
        elif tag == 'Categories':
            for cat in child._children:
                self.categories.add(cat.text)
        elif tag == 'Attachments':
            self.attachments.populate_from_node(self.service, child)
        elif tag in self.boolean_fields_tag or tag.startswith('Is'):
            # boolean
            setattr(self.mapping_dict_tag_obj[tag],
                    'value',
                    child.text == 'true')
        else:
            return False
        return True

    def _init_base_fields_from_resp(self, rnode):
        """Return a reference to the parsed Element object for response after
        snarfing all the common fields."""

        for child in rnode:
            tag = unQName(child.tag)
            # if tag == 'ItemId':
            #     import pdb; pdb.set_trace()
            if tag in self.mapping_dict_tag_obj:
                res = self._process_tag(child, tag)
                if not res:
                    setattr(self.mapping_dict_tag_obj[tag],
                            'value',
                            child.text)

            elif tag == 'ExtendedProperty':
                self.add_extended_property(node=child)
            else:
                _logger.warning(
                    'Trying to instanciate an unknown field %s' % tag)

# <Item>
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
#    <Importance/>
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
#    <EffectiveRights/>
#    <LastModifiedName/>
#    <LastModifiedTime/>
#    <IsAssociated/>
#    <WebClientReadFormQueryString/>
#    <WebClientEditFormQueryString/>
#    <ConversationId/>
#    <UniqueBody/>
# </Item>
