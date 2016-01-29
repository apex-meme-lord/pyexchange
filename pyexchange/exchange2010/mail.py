import logging
from lxml import etree

from ..base.message import (
  BaseExchangeMessageService, BaseExchangeMessage, BaseExchangeMessageList,
  ExchangeMailboxTarget, ExchangeMailboxTargetList
)

from ..utils import auto_build_dict_from_xml

from . import soap_request

log = logging.getLogger('pyexchange')


class Exchange2010MessageService(BaseExchangeMessageService):

  def list_messages(self, folder_id, delegate_for=None):
    return Exchange2010MessageList(service=self.service, folder_id=folder_id, delegate_for=delegate_for)

  def list_messages_batch(self, folder_id, max_entries=100, delegate_for=None):
    return Exchange2010MessageGenerator(service=self.service, folder_id=folder_id, max_entries=max_entries, delegate_for=delegate_for)

  def get_message(self, id):
    return Exchange2010Message(service=self.service, id=id)

  def new_message(self, **kwargs):
    return Exchange2010Message(service=self.service, **kwargs)


class Exchange2010Message(BaseExchangeMessage):
  
  def _init_from_service(self, id):
    log.debug(u'Creating new Exchange2010CalendarEvent object from ID')
    
    request = soap_request.get_message(exchange_id=id, format=u'AllProperties')
    response = self.service.send(request)
    
    items = response.xpath(u'//m:GetItemResponseMessage/m:Items/t:Message', namespaces=soap_request.NAMESPACES)  

    body_content, body_type = self._parse_body_content_and_type(items[0])
    self.body = body_content
    self.body.type = body_type

    return self._init_from_xml(items[0])

  def _init_from_xml(self, xml):
    log.debug(u'Creating new Exchange2010 object from XML')

    self._id, self._change_key = self._parse_id_and_change_key(xml)
    self._parent_folder_id, self._parent_folder_change_key = self._parse_parent_id_and_change_key(xml)

    properties = self._parse_response_for_other_props(xml)
    self._update_properties(properties)

    log.debug(u'Created new message object with ID: %s' % self._id)
    self._reset_dirty_attributes()

    return self

  def _parse_id_and_change_key(self, xml):
    id_elements = xml.xpath(u'//t:Message/t:ItemId', namespaces=soap_request.NAMESPACES)

    if id_elements:
      id_element = id_elements[0]
      id_element.getparent().remove(id_element)
      return id_element.get(u'Id', None), id_element.get(u'ChangeKey', None)
    return None, None

  def _parse_parent_id_and_change_key(self, xml):
    parent_ids = xml.xpath(u'//t:Message/t:ParentFolderId', namespaces=soap_request.NAMESPACES)

    if parent_ids:
      parent_id = parent_ids[0]
      parent_id.getparent().remove(parent_id)
      return parent_id.get(u'Id', None), parent_id.get(u'ChangeKey', None)
    return None, None

  def _parse_body_content_and_type(self, xml):
    body_elements = xml.xpath(u'//t:Message/t:Body', namespaces=soap_request.NAMESPACES)

    if body_elements:
      body = body_elements[0]
      body.getparent().remove(body)
      return body.text, body.get(u'BodyType')
    return None, None


  def _parse_response_for_other_props(self, xml):

    property_map = auto_build_dict_from_xml(xml)

    # this is tedious, but is a lot simpler than a deep merge of two dicts
    typecast_map = {
      u'date_time_received': u'datetime',
      u'size': u'int',
      u'is_submitted': u'bool',
      u'is_draft': u'bool',
      u'is_from_me': u'bool',
      u'is_resend': u'bool',
      u'is_unmodified': u'bool',
      u'date_time_sent': u'datetime',
      u'date_time_created': u'datetime',
      u'has_attachments': u'bool',
      u'is_read_receipt_requested': u'bool',
      u'is_read': u'bool'
    }

    for prop in property_map:
      if prop in typecast_map:
        property_map[prop][u'cast'] = typecast_map[prop]

    return self.service._xpath_to_dict(element=xml, property_map=property_map, namespace_map=soap_request.NAMESPACES)

  def _fetch_message_body(self):
    body = soap_request.get_message(format=u'AllProperties', exchange_id=self.id)
    response = self.service.send(body)

  def create(self):
    pass

  def update(self):
    pass

  def delete(self):
    pass

  def validate(self):
    pass

  def send(self):
    pass

  def copy(self, folder_id):
    pass

  def move(self, folder_id):
    pass


class Exchange2010MessageList(BaseExchangeMessageList):

  def _fetch_message_items(self, folder_id, delegate_for):
    request = soap_request.get_message_items(format=u'AllProperties', folder_id=folder_id, delegate_for=delegate_for)
    return self.service.send(request)
  
  def _parse_response_for_list_or_get_messages(self, response):
    items = (
      response.xpath(u'//m:FindItemResponseMessage/m:RootFolder/t:Items/t:Message', namespaces=soap_request.NAMESPACES)
      or response.xpath(u'//m:FindItemResponseMessage/m:RootFolder/t:Groups/t:GroupedItems/t:Items/t:Message', namespaces=soap_request.NAMESPACES)
      or None
    )

    for item in items:
      self._add_message(item)

    return self

  def _add_message(self, xml):
    self._messages.append(Exchange2010Message(service=self.service, xml=xml))
    return self
  

class Exchange2010MessageGenerator(Exchange2010MessageList):

  def _fetch_messages_items(self, folder_id, delegate_for, max_entries):
    request = soap_request.get_message_items(format=u'AllProperties', folder_id=folder_id, delegate_for=delegate_for, max_entries=max_entries)
    return self.service.send(request)

  def _parse_response_for_messages(self, response):
    return super(Exchange2010MessageGenerator, self)._parse_response_for_messages(response)