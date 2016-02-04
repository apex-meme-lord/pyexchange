import logging
from lxml import etree

from ..base.message import (
  BaseExchangeMessageService, BaseExchangeMessage, BaseExchangeMessageList,
  ExchangeMailboxTarget, ExchangeMailboxTargetList,
  ExchangeAttachment, ExchangeAttachmentList,
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


class Exchange2010MailboxTarget(ExchangeMailboxTarget):

  def _init_from_xml(self, xml):
    parsed = self._parse_info_from_mailbox_xml(xml)
    self._init_from_props(**parsed)

  def _init_from_props(self, name, email_address, routing_type=u'SMTP'):
    self.name = name
    self.email_address = email_address
    self.routing_type = routing_type

  def _parse_info_from_mailbox_xml(self, xml):
    
    property_map = {
      u'name': {
        u'xpath': u'./t:Name'
      },
      u'email_address': {
        u'xpath': u'./t:EmailAddress'
      },
      u'routing_type': {
        u'xpath': u'./t:RoutingType'
      }
    }
    return self.service._xpath_to_dict(element=xml, property_map=property_map, namespace_map=soap_request.NAMESPACES)


class Exchange2010MailboxTargetList(ExchangeMailboxTargetList):

  def _parse_mailbox_from_xml(self, xml):
    for mailbox in xml.getchildren():
      self._add_mailbox(xml=mailbox)

    return self

  def _add_mailbox(self, **kwargs):
    new_mailbox = Exchange2010MailboxTarget(service=self.service, **kwargs)
    self._mailboxes.append(new_mailbox)
    return new_mailbox


class Exchange2010Attachment(ExchangeAttachment):
  
  def _init_from_xml(self, xml):
    properties = self._parse_info_from_attachment_xml(xml)
    self._update_properties(properties)

  def _init_from_service(self, id):
    xml = self._fetch_from_service(id)
    properties = self._parse_info_from_attachment_xml(xml)
    self._update_properties(properties)

  def _update_properties(self, properties):
    for key in properties:
      setattr(self, key, properties[key])

  def _fetch_from_service(self, id):
    request = soap_request.get_attachment(id)
    response = self.service.send(request)
    return response.xpath('//m:GetAttachmentResponseMessage/m:Attachments/t:FileAttachment', namespaces=soap_request.NAMESPACES)[0]

  def _parse_info_from_attachment_xml(self, xml):
    property_map = {
      u'name': {
        u'xpath': u'./t:Name',
      },
      u'content_type': {
        u'xpath': u'./t:ContentType',
      },
      u'content': {
        u'xpath': u'./t:Content',
      }
    }    
    properties = self.service._xpath_to_dict(element=xml, property_map=property_map, namespace_map=soap_request.NAMESPACES)
    id = xml.xpath('./t:AttachmentId', namespaces=soap_request.NAMESPACES)

    if len(id):
      properties.update(id=id[0].attrib['Id'])

    return properties

  @property
  def content_type(self):
    if self._content_type is None and self.id:
      xml = self._fetch_from_service(self.id)
      self._content_type = self._parse_info_from_attachment_xml(xml)['content_type']

    return self._content_type

  @content_type.setter
  def content_type(self, value):
    self._content_type = value

  @property
  def content(self):
    if self._content is None and self.id:
      xml = self._fetch_from_service(self.id)
      self._content = self._parse_info_from_attachment_xml(xml)['content']

    return self._content

  @content.setter
  def content(self, value):
    self._content = value


class Exchange2010AttachmentList(ExchangeAttachmentList):
  def _parse_attachments_from_xml(self, xml):
    for attachment in xml.getchildren():
      self._add_attachment(xml=attachment)

  def _add_attachment(self, xml):
    new_attachment = Exchange2010Attachment(service=self.service, xml=xml)
    self._attachments.append(new_attachment)
    return new_attachment
    


class Exchange2010Message(BaseExchangeMessage):
  
  def _init_from_service(self, id):
    log.debug(u'Creating new Exchange2010CalendarEvent object from ID')
  
    message = self._fetch_message_from_service(id)

    self.body = self._parse_body_content_and_type(message)
    self.to_recipients = self._parse_to_recipients(message)
    self.cc_recipients = self._parse_cc_recipients(message)
    self.reply_to = self._parse_reply_to(message)
    self.attachments = self._parse_attachment(xml)

    return self._init_from_xml(message)

  def _init_from_xml(self, xml):
    log.debug(u'Creating new Exchange2010 object from XML')

    self._id, self._change_key = self._parse_id_and_change_key(xml)
    self._parent_folder_id, self._parent_folder_change_key = self._parse_parent_id_and_change_key(xml)
    self.sender = self._parse_sender(xml)
    self.from_ = self._parse_from(xml)

    properties = self._parse_response_for_other_props(xml)
    self._update_properties(properties)

    log.debug(u'Created new message object with ID: %s' % self._id)
    self._reset_dirty_attributes()

    return self

  def _init_from_mime_content(self, mime, folder_id, character_set):
    log.debug(u'Creating new Exchange2010 object from MIME content')

    request = soap_request.new_message_from_mime(mime_content=mime, folder_id=folder_id, character_set=character_set)
    response = self.service.send(request)

    return self._init_from_xml(response)

  def _pop_element(self, xml, xpath):
    element = xml.xpath(xpath, namespaces=soap_request.NAMESPACES)

    if len(element):
      exists = element[0]
      exists.getparent().remove(exists)
      return exists

  def _parse_id_and_change_key(self, xml):
    item_id = self._pop_element(xml, u'./t:ItemId')
    if item_id is not None:
      return item_id.get(u'Id', None), item_id.get(u'ChangeKey', None)
 
    return None, None

  def _parse_parent_id_and_change_key(self, xml):
    parent_folder_id = self._pop_element(xml, u'./t:ParentFolderId')
    if parent_folder_id is not None:
      return parent_folder_id.get(u'Id', None), parent_folder_id.get(u'ChangeKey', None)
 
    return None, None

  def _parse_body_content_and_type(self, xml):
    self.hehehe = xml
    return self._pop_element(xml, u'./t:Body')

  def _parse_to_recipients(self, xml):
    return self._pop_element(xml, u'./t:ToRecipients')
    
  def _parse_cc_recipients(self, xml):
    return self._pop_element(xml, u'./t:CcRecipients')
  
  def _parse_reply_to(self, xml):
    return self._pop_element(xml, u'./t:ReplyTo')

  def _parse_sender(self, xml):
    return self._pop_element(xml, u'./t:Sender')

  def _parse_from(self, xml):
    return self._pop_element(xml, u'./t:From')

  def _parse_attachments(self, xml):
    return self._pop_element(xml, u'./t:Attachments')

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

  def create(self):
    request = soap_request.new_message_save_only(message=self)
    response = self.service.send(request)

    for attachment in self.attachments:
      self.create_attachment(name=attachment.name, content=attachment.content)

    self.refresh_id_and_change_key()
    return self

  def delete(self):
    request = soap_request.delete_message(message=self)
    response = self.service.send(request)

  def validate(self):
    pass

  def send(self):
    request = soap_request.copy_message(message=self)
    response = self.service.send(request)

    return self

  def copy(self, folder_id):
    request = soap_request.copy_message(message=self, folder_id=folder_id)
    response = self.service.send(request)

    return self

  def move(self, folder_id):
    request = soap_request.move_item(id=self, folder_id=folder_id)
    response = self.service.send(request)

    self.refresh_parent_folder()
    self.refresh_id_and_change_key()

    return self

  def create_attachment(self, name, content):
    request = soap_request.new_attachment(item=self, name=name, content=content)
    response = self.service.send(request)
    id= response.xpath('//m:CreateAttachmentResponseMessage/m:Attachments/t:FileAttachment/t:AttachmentId', namespaces=soap_request.NAMESPACES)[0].attrib['Id']

    self.refresh_id_and_change_key()
    
    return Exchange2010Attachment(service=self.service, id=id)

  @property
  def to_recipients(self):
    if self._to_recipients is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.to_recipients = self._parse_to_recipients(message)
 
    return self._to_recipients

  @to_recipients.setter
  def to_recipients(self, value):
      self._to_recipients = Exchange2010MailboxTargetList(service=self.service, xml=value)

  @property
  def cc_recipients(self):
    if self._cc_recipients is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.cc_recipients = self._parse_cc_recipients(message)

    return self._cc_recipients

  @cc_recipients.setter
  def cc_recipients(self, value):
      self._cc_recipients = Exchange2010MailboxTargetList(service=self.service, xml=value)

  @property
  def reply_to(self):
    if self._reply_to is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.reply_to = self._parse_reply_to(message)

    return self._reply_to

  @reply_to.setter
  def reply_to(self, value):
      self._reply_to = Exchange2010MailboxTargetList(service=self.service, xml=value)

  @property
  def sender(self):
    if self._sender is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.sender = self._parse_sender(message)

    return self._sender

  @sender.setter
  def sender(self, value):
      self._sender = Exchange2010MailboxTargetList(service=self.service, xml=value)

  @property
  def from_(self):
    if self._from_ is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.from_ = self._parse_from(message)

    return self._from_

  @from_.setter
  def from_(self, value):
      self._from_ = Exchange2010MailboxTargetList(service=self.service, xml=value)

  @property
  def attachments(self):
    if self._attachments is None and self.id:
      message = self._fetch_message_from_service(self.id)
      self.attachments = self._parse_attachments(message)

    return self._attachments
  
  @attachments.setter
  def attachments(self, value):
    self._attachments = Exchange2010AttachmentList(service=self.service, xml=value)

  def refresh_id_and_change_key(self):
    message = self._fetch_message_from_service(self.id)
    self._id, self._change_key = self._parse_id_and_change_key(message)

    return self

  def refresh_parent_folder(self):
    message = self._fetch_message_from_service(self.id)
    self._parent_folder_id, self._parent_folder_change_key = self._parse_parent_id_and_change_key(message)

    return self

  def _fetch_message_from_service(self, id):
    request = soap_request.get_message(exchange_id=id, format=u'AllProperties')
    response = self.service.send(request)
    
    items = response.xpath(u'//m:GetItemResponseMessage/m:Items/t:Message', namespaces=soap_request.NAMESPACES)  
    
    if items:
      return items[0]

    raise RuntimeError('Unable to fetch message from id: {}'.format(id))

  def _fetch_message_body(self):
    message = self._fetch_message_from_service(self.id)
    return self._parse_body_content_and_type(message)

  def _fetch_attachments(self):
    message = self._fetch_message_from_service(self.id)
    return self._parse_attachment(message)


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
  
  def send(self):
    request = soap_request.send_messages(messages=self._messages)
    response = self.service.send(request)

    return self

  def copy(self, folder_id):
    request = soap_request.copy_messages(messages=self._messages, folder_id=folder_id)
    response = self.service.send(request)

    return self.service.mail().list_messages(folder_id=folder_id)


class Exchange2010MessageGenerator(Exchange2010MessageList):

  def _fetch_messages_items(self, folder_id, delegate_for, max_entries):
    request = soap_request.get_message_items(format=u'AllProperties', folder_id=folder_id, delegate_for=delegate_for, max_entries=max_entries)
    return self.service.send(request)

  def _parse_response_for_messages(self, response):
    return super(Exchange2010MessageGenerator, self)._parse_response_for_messages(response)