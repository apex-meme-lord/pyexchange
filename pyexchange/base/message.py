class ExchangeMessageBody(object):

  def __init__(self, text=None, html=None, xml=None):
    if text is not None:
      self.content = text
      self._type = u'Text'

    elif html is not None:
      self.content = html
      self._type = u'HTML'

    elif xml is not None:
      self.content = xml.text
      self._type = xml.attrib['BodyType']

  @property
  def type(self):
      return self._type


class ExchangeMailboxTargetList(object):

  def __init__(self, service, xml=None):
    self.service = service
    self._mailboxes = []
    
    if xml is not None:
      self._parse_mailbox_from_xml(xml)

  def _parse_mailbox_from_xml(self, xml):
    raise NotImplementedError

  def _add_mailbox(self, **kwargs):
    raise NotImplementedError

  def __repr__(self):
    return repr(self._mailboxes)

  def __len__(self):
    return len(self._mailboxes)

  def __getitem__(self, key):
    return self._mailboxes[key]

  def __setitem__(self, key, value):
    self._mailboxes[key] = value

  def __delitem__(self, key):
    del self._mailboxes[key]

  def __iter__(self):
    return self._mailboxes.__iter__()

  def __reversed__(self):
    return self._mailboxes.__reversed__()

  def __contains__(self, item):
    return self._mailboxes.__contains__(item)

  def append(self, *args, **kwargs):
    return self._messages._add_mailbox(*args, **kwargs)

  def pop(self, *args, **kwargs):
    return self._messages.pop(*args, **kwargs)


class ExchangeMailboxTarget(object):

  name = None
  email_address = None
  routing_type = None
  mailbox_type = None

  _type = None

  def __init__(self, service, xml=None, **kwargs):
    self.service = service

    if xml is not None:
      self._init_from_xml(xml)
    else:
      self._init_from_props(**kwargs)

  def _init_from_xml(self, xml):
    raise NotImplementedError

  def _init_from_props(self, **kwargs):
    raise NotImplementedError


class ExchangeAttachmentList(object):
  
  def __init__(self, service, xml=None):
    self.service = service
    self._attachments = []

    if xml is not None:
      self._parse_attachments_from_xml(xml)

  def _parse_attachments_from_xml(self, xml):
    raise NotImplementedError

  def _add_attachment(self, **kwargs):
    raise NotImplementedError

  def __repr__(self):
    return repr(self._attachments)

  def __len__(self):
    return len(self._attachments)

  def __getitem__(self, key):
    return self._attachments[key]

  def __setitem__(self, key, value):
    self._attachments[key] = value

  def __delitem__(self, key):
    del self._attachments[key]

  def __iter__(self):
    return self._attachments.__iter__()

  def __reversed__(self):
    return self._attachments.__reversed__()

  def __contains__(self, item):
    return self._attachments.__contains__(item)

  def append(self, *args, **kwargs):
    return self._attachments._add_attachment(*args, **kwargs)

  def pop(self, *args, **kwargs):
    return self._attachments.pop(*args, **kwargs)


class ExchangeAttachment(object):

  _content = None
  _content_type = None

  def __init__(self, service, id=None, xml=None):
    self.service = service

    if xml is not None:
      self._init_from_xml(xml)
    elif id is not None:
      self._init_from_service(id)

  def _init_from_xml(self, xml):
    raise NotImplementedError

  def _init_from_service(self, id):
    raise NotImplementedError

  def _init_from_filepath(self, name, filepath):
    raise NotImplementedError

  def _init_from_props(self, name, content):
    raise NotImplementedError


class BaseExchangeMessageService(object):

  def __init__(self, service):
    self.service = service

  def list_messages(self, folder_id):
    raise NotImplementedError

  def get_message(self, id):
    raise NotImplementedError

  def new_message(self, **kwargs):
    raise NotImplementedError


class BaseExchangeMessageList(object):

  def __init__(self, service, folder_id, delegate_for=None, **kwargs):
    
    self.service = service
    self._messages = []

    response = self._fetch_message_items(folder_id, delegate_for, **kwargs)
    self._parse_response_for_list_or_get_messages(response)

  def _fetch_message_items(self, folder_id, delegate_for):
    raise NotImplementedError
  
  def _parse_response_for_list_or_get_messages(self, response):
    raise NotImplementedError

  def __repr__(self):
    return repr(self._messages)

  def __len__(self):
    return len(self._messages)

  def __getitem__(self, key):
    return self._messages[key]

  def __setitem__(self, key, value):
    self._messages[key] = value

  def __delitem__(self, key):
    del self._messages[key]

  def __iter__(self):
    return self._messages.__iter__()

  def __reversed__(self):
    return self._messages.__reversed__()

  def __contains__(self, item):
    return self._messages.__contains__(item)

  def append(self, *args, **kwargs):
    return self._messages._add_message(*args, **kwargs)

  def pop(self, *args, **kwargs):
    return self._messages.pop(*args, **kwargs)


class BaseExchangeMessage(object):

  _id = None
  _change_key = None

  _parent_folder_id = None
  _parent_folder_change_key = None

  _item_class = None
  
  subject = None
  sensitivity = None
  date_time_received = None
  size = None
  importance = None
  is_submitted = None
  is_draft = None
  is_from_me = None
  is_resend = None
  is_unmodified = None
  date_time_sent = None
  date_time_created = None
  display_cc = None
  display_to = None
  has_attachments = None
  culture = None
  is_read_receipt_requested = None
  conversation_index = None
  conversation_topic = None
  internet_message_id = None
  is_read = None

  _response_objects = None
  
  _to_recipients = None
  _cc_recipients = None
  _reply_to = None
  _from_ = None
  _sender = None

  _attachments = None
  _internet_message_headers = None
  
  _body = None

  _track_dirty_attributes = False
  _dirty_attributes = set()

  def __init__(self, service, id=None, xml=None, mime=None, folder_id=None, character_set=u'UTF-8', **kwargs):

    self.service = service

    if xml is not None:
      self._init_from_xml(xml)
    elif id is not None:
      self._init_from_service(id)
    elif mime is not None:
      if folder_id is None:
        raise ValueError('Creating an Exchange message from MIME requires a folder id and character set encoding')
      self._init_from_mime_content(mime, folder_id, character_set)
    else:
      self._update_properties(kwargs)


  def _init_from_xml(self, xml):
    raise NotImplementedError

  def _init_from_service(self, id):
    raise NotImplementedError

  def _init_from_mime_content(self, mime):
    raise NotImplementedError

  def create(self):
    raise NotImplementedError

  def update(self):
    raise NotImplementedError

  def delete(self):
    raise NotImplementedError

  def validate(self):
    raise NotImplementedError

  def send(self):
    raise NotImplementedError

  def copy(self, folder_id):
    raise NotImplementedError

  def move(self, folder_id):
    raise NotImplementedError

  @property
  def id(self):
    """ **Read-only.** The internal id Exchange uses to refer to this message. """
    return self._id

  @property
  def change_key(self):
    """ **Read-only.** When you change an message, Exchange makes you pass a change key to prevent overwriting a previous version. """
    return self._change_key

  @property
  def parent_folder_id(self):
    return self._parent_folder_id
    
  @parent_folder_id.setter
  def parent_folder_id(self, value):
    self._parent_folder_id = value
  
  @property
  def parent_folder_change_key(self):
    return self._parent_folder_change_key
    
  @parent_folder_change_key.setter
  def parent_folder_change_key(self, value):
    self._parent_folder_change_key = value
  
  @property
  def item_class(self):
    return self._item_class

  @item_class.setter
  def item_class(self, value):
    if self._item_class is None:
      self._item_class = value

  @property
  def body(self):
    """Lazy-loaded message body"""
    if self._body is None and self.id is not None:
      self.body = self._fetch_message_body()
    return self._body

  @body.setter
  def body(self, value):
    if isinstance(value, str):
      self._body = ExchangeMessageBody(text=value)
    elif hasattr(value, 'text') and hasattr(value, 'attrib'):
      # duck-typing magic
      self._body = ExchangeMessageBody(xml=value)
    else:
      self._body = ExchangeMessageBody(text=u'')

  def add_file_attachment(self, fully_qualified_filepath=None, new_filename=None, byte_array=None, input_stream=None):
    raise NotImplementedError

  def add_html_attachment(self):
    raise NotImplementedError

  def _update_properties(self, properties):
    self._track_dirty_attributes = False
    for key in properties:
      setattr(self, key, properties[key])
    self._track_dirty_attributes = True

  def __setattr__(self, key, value):
    """ Magically track public attributes, so we can track what we need to flush to the Exchange store """
    if self._track_dirty_attributes and not key.startswith(u"_"):
      self._dirty_attributes.add(key)

    try:
      super(BaseExchangeMessage, self).__setattr__(key, value)
    except:
      raise AttributeError('%s could not be set to %r' % (key, value))

  def _reset_dirty_attributes(self):
    self._dirty_attributes = set()

  def _fetch_message_to_recipients(self):
    raise NotImplementedError

  def _fetch_message_body(self):
    raise NotImplementedError