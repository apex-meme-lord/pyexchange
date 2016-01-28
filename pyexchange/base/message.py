class ExchangeMessageBody(object):

  def __init__(self, content, type_=None):
    self.content = content
    self.type = type_


class ExchangeMailboxTarget(object):

  name = None
  email_address = None
  routing_type = None

  _type = None

  def __init__(self, xml=None, props=None):
    if xml:
      self._init_from_xml(xml)
    if props:
      self._init_from_props(props)

  def _init_from_xml(self, xml):
    raise NotImplementedError

  def _init_from_props(self, props):
    raise NotImplementedError


class ExchangeMailboxTargetList(object):

  def __init__(self, xml):

    self._mailboxes = []
    
    for mailbox in xml.getchildren():
      self._parse_mailbox_from_xml(mailbox)

  def _parse_mailbox_from_xml(self, xml):
    raise NotImplementedError

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


class BaseExchangeMessageService(object):

  def __init__(self, service):
    self.service = service

  def list_messages(self, folder_id):
    raise NotImplementedError

  def get_message(self, id):
    raise NotImplementedError

  def create_message(self, **kwargs):
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
    return self._messages.append(*args, **kwargs)

  def insert(self, *args, **kwargs):
    return self._messages.insert(*args, **kwargs)

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
  sender = None
  is_read_receipt_requested = None
  conversation_index = None
  conversation_topic = None
  from_ = None
  internet_message_id = None
  is_read = None

  _response_objects = None
  _to_recipients = None
  _attachments = None
  _internet_message_headers = None
  _body = None

  _track_dirty_attributes = False
  _dirty_attributes = set()

  def __init__(self, service, id=None, xml=None, **kwargs):

    self.service = service

    if xml is not None:
      self._init_from_xml(xml)
    elif id is None:
      self._update_properties(kwargs)
    else:
      self._init_from_service(id)

  def _init_from_xml(self, xml):
    raise NotImplementedError

  def _init_from_service(self, id):
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

  @id.setter
  def id(self, value):
    self._id = value

  @property
  def change_key(self):
    """ **Read-only.** When you change an message, Exchange makes you pass a change key to prevent overwriting a previous version. """
    return self._change_key

  @change_key.setter
  def change_key(self, value):
    self._change_key = value

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
    self._item_class = value
  
  @property
  def body(self):
    return self._body

  @body.setter
  def body(self, value):
    self._body = ExchangeMessageBody(value, u'Text')

  @property
  def to_recipients(self):
    return self._to_recipients

  @to_recipients.setter
  def to_recipients(self, value):
    self._to_recipients = value 

  

  # @property
  # def mime_content(self):
  #   return self._mime_content

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