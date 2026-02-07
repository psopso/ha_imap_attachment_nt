"""Email attachment sensor support."""
from collections import deque
import datetime
import email
import imaplib
import logging
import os
import pandas as pd

import voluptuous as vol

from homeassistant.components.sensor import PLATFORM_SCHEMA
from homeassistant.const import (
    ATTR_DATE,
    CONF_NAME,
    CONF_PASSWORD,
    CONF_PORT,
    CONF_USERNAME,
    CONF_VALUE_TEMPLATE,
    CONTENT_TYPE_TEXT_PLAIN,
)
import homeassistant.helpers.config_validation as cv
from homeassistant.helpers.entity import Entity

import unicodedata
import re
from email.header import decode_header, make_header

_LOGGER = logging.getLogger(__name__)

CONF_SERVER = "server"
CONF_SENDERS = "senders"
CONF_FOLDER = "folder"
CONF_STORAGE_PATH = "storage_path"
CONF_CSV_FILENAME = "csv_filename"

ATTR_FROM = "from"
ATTR_BODY = "body"
ATTR_SUBJECT = "subject"
ATTR_NUM_ATTACHMENTS = "num_attachments"
ATTR_ATTACHMENT_PATHS = "attachment_paths"

DEFAULT_PORT = 993

PLATFORM_SCHEMA = PLATFORM_SCHEMA.extend(
    {
        vol.Optional(CONF_NAME): cv.string,
        vol.Required(CONF_USERNAME): cv.string,
        vol.Required(CONF_PASSWORD): cv.string,
        vol.Required(CONF_SERVER): cv.string,
        vol.Required(CONF_SENDERS): [cv.string],
        vol.Optional(CONF_PORT, default=DEFAULT_PORT): cv.port,
        vol.Optional(CONF_VALUE_TEMPLATE): cv.template,
        vol.Optional(CONF_FOLDER, default="INBOX"): cv.string,
        vol.Optional(CONF_STORAGE_PATH, default="/config/attachments"): cv.string,
        vol.Optional(CONF_CSV_FILENAME, default="tariff.save"): cv.string,
    }
)


def setup_platform(hass, config, add_entities, discovery_info=None):
    """Set up the Email sensor platform."""
    reader = EmailReader(
        config.get(CONF_USERNAME),
        config.get(CONF_PASSWORD),
        config.get(CONF_SERVER),
        config.get(CONF_PORT),
        config.get(CONF_FOLDER),
        config.get(CONF_STORAGE_PATH),
        config.get(CONF_CSV_FILENAME),
    )

    storage_path = config.get(CONF_STORAGE_PATH)
    if not os.path.exists(storage_path):
        os.makedirs(storage_path)

    csv_filename = config.get(CONF_CSV_FILENAME)

    value_template = config.get(CONF_VALUE_TEMPLATE)
    if value_template is not None:
        value_template.hass = hass
    sensor = EmailContentSensor(
        hass,
        reader,
        config.get(CONF_NAME) or config.get(CONF_USERNAME),
        config.get(CONF_SENDERS),
        value_template,
        config.get(CONF_STORAGE_PATH),
        config.get(CONF_CSV_FILENAME),
    )

    if sensor.connected:
        add_entities([sensor], True)
    else:
        return False


class EmailReader:
    """A class to read emails from an IMAP server."""

    def __init__(self, user, password, server, port, folder, storage_path, csv_filename):
        """Initialize the Email Reader."""
        self._user = user
        self._password = password
        self._server = server
        self._port = port
        self._folder = folder
        self._last_id = None
        self._unread_ids = deque([])
        self._storage_path = storage_path
        self._csv_filename = csv_filename
        self.connection = None
        self._message_data = None
        self._last_uid = None

    def connect(self):
        """Login and setup the connection."""
        try:
            self.connection = imaplib.IMAP4_SSL(self._server, self._port)
            self.connection.login(self._user, self._password)
            return True
        except imaplib.IMAP4.error:
            _LOGGER.error("Failed to login to %s", self._server)
            return False

    def _fetch_message(self, message_uid):
        """Get an email message from a message id."""
        _, message_data = self.connection.uid("fetch", message_uid, "(RFC822)")

        if message_data is None:
            return None
        if message_data[0] is None:
            return None

        self._message_data = message_data
        #_LOGGER.info(message_data)

        raw_email = message_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        return email_message

    def read_next(self):
        """Read the next email from the email server."""
        try:
            message_uid = None
            self.connection.select(self._folder, readonly=True)
            _LOGGER.info('read_next')
            if not self._unread_ids:
                search = f'(SINCE {datetime.date.today():%d-%b-%Y} SUBJECT "Export" SUBJECT "NT")'
                if self._last_id is not None:
                    search = f'(UID {self._last_id}:* SUBJECT "Export")'

                _, data = self.connection.uid("search", None, search)
                self._unread_ids = deque(data[0].split())

            while self._unread_ids:
                message_uid = self._unread_ids.popleft()
                _LOGGER.info('messageuid:')
                if self._last_id is None or int(message_uid) > self._last_id:
                    self._last_id = int(message_uid)
                    self._last_uid = message_uid
                    _LOGGER.info('Next message: ' + message_uid.decode("ascii"))
                    #self.connection.uid("COPY", message_uid, '"[Gmail]/Trash"')
                    #self.connection.uid("STORE", message_uid, "+FLAGS", r"(\Deleted)")
                    return self._fetch_message(message_uid)
            
            if message_uid is not None:
                self._last_uid = message_uid
                #_LOGGER.info('Next message: ' + message_uid.decode("ascii"))
                #self.connection.uid("COPY", message_uid, '"[Gmail]/Trash"')
                #self.connection.uid("STORE", message_uid, "+FLAGS", r"(\Deleted)")
                #_LOGGER.info('Message deleted?: ' + message_uid.decode("ascii"))

                status1, data1 = self.connection.uid("fetch", message_uid, '(FLAGS)')
                _LOGGER.info(data1);

            self.connection.expunge()

            return self._fetch_message(str(self._last_id))

        except imaplib.IMAP4.error:
            _LOGGER.info("Connection to %s lost, attempting to reconnect", self._server)
            try:
                self.connect()
                _LOGGER.info(
                    "Reconnect to %s succeeded, trying last message", self._server
                )
                if self._last_id is not None:
                    return self._fetch_message(str(self._last_id))
            except imaplib.IMAP4.error:
                _LOGGER.error("Failed to reconnect")

        return None


class EmailContentSensor(Entity):
    """Representation of an EMail sensor."""

    def __init__(self, hass, email_reader, name, allowed_senders, value_template, storage_path, csv_filename):
        """Initialize the sensor."""
        self.hass = hass
        self._email_reader = email_reader
        self._name = name
        self._allowed_senders = [sender.upper() for sender in allowed_senders]
        self._value_template = value_template
        self._storage_path = storage_path
        self._csv_filename = csv_filename
        self._last_id = None
        self._message = None
        self._state_attributes = None
        self.connected = self._email_reader.connect()

    @property
    def name(self):
        """Return the name of the sensor."""
        return self._name

    @property
    def state(self):
        """Return the current email state."""
        return self._message

    @property
    def device_state_attributes(self):
        """Return other state attributes for the message."""
        return self._state_attributes

    def render_template(self, email_message):
        _LOGGER.info('rendertemplate')
        """Render the message template."""
        variables = {
            ATTR_FROM: EmailContentSensor.get_msg_sender(email_message),
            ATTR_SUBJECT: EmailContentSensor.get_msg_subject(email_message),
            ATTR_DATE: email_message["Date"],
            ATTR_BODY: EmailContentSensor.get_msg_text(email_message),
            ATTR_NUM_ATTACHMENTS: EmailContentSensor.get_num_msg_attachments(email_message),
            ATTR_ATTACHMENT_PATHS: EmailContentSensor.get_msg_attachments(email_message, self._storage_path, self._csv_filename),
        }
        return self._value_template.render(variables, parse_result=False)

    def sender_allowed(self, email_message):
        """Check if the sender is in the allowed senders list."""
        return EmailContentSensor.get_msg_sender(email_message).upper() in (
            sender for sender in self._allowed_senders
        )

    @staticmethod
    def get_msg_sender(email_message):
        """Get the parsed message sender from the email."""
        return str(email.utils.parseaddr(email_message["From"])[1])

    @staticmethod
    def get_msg_subject(email_message):
        """Decode the message subject."""
        decoded_header = email.header.decode_header(email_message["Subject"])
        header = email.header.make_header(decoded_header)
        return str(header)

    @staticmethod
    def get_msg_text(email_message):
        """
        Get the message text from the email.
        Will look for text/plain or use text/html if not found.
        """
        message_text = None
        message_html = None
        message_untyped_text = None

        for part in email_message.walk():
            if part.get_content_type() == CONTENT_TYPE_TEXT_PLAIN:
                if message_text is None:
                    message_text = part.get_payload()
            elif part.get_content_type() == "text/html":
                if message_html is None:
                    message_html = part.get_payload()
            elif part.get_content_type().startswith("text"):
                if message_untyped_text is None:
                    message_untyped_text = part.get_payload()

        if message_text is not None:
            return message_text

        if message_html is not None:
            return message_html

        if message_untyped_text is not None:
            return message_untyped_text

        return email_message.get_payload()

    @staticmethod
    def get_num_msg_attachments(email_message):
        """
        Detect number of attachments.
        First index is the email body.
        """
        return len(email_message.get_payload()) - 1

    @staticmethod
    def get_msg_attachments(email_message, storage_path, csv_filename):
        """
        Parse attachments on the email, clean filenames and store locally.
        """
        attachments = []
        
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            filename = part.get_filename()

            if filename:
                # 1. DEKÓDOVÁNÍ (z MIME formátu =?utf-8?Q?...)
                try:
                    # decode_header vrací seznam dvojic (bytes, encoding)
                    decoded_header = decode_header(filename)
                    # make_header to spojí do jednoho objektu, str() to převede na čitelný řetězec
                    filename = str(make_header(decoded_header))
                except Exception as e:
                    _LOGGER.warning(f"Error decoding header: {e}")
                    # Pokud selže, necháme původní název

                # 2. ODSTRANĚNÍ ČEŠTINY A VYČIŠTĚNÍ
                # Normalizace rozdělí znaky (např. 'č' na 'c' + háček)
                normalized = unicodedata.normalize('NFKD', filename)
                # Encode do ASCII zahodí (ignore) znaky, co nejsou ASCII (tím zmizí háčky/čárky)
                filename_ascii = normalized.encode('ASCII', 'ignore').decode('ASCII')
                # Regulární výraz: povolí jen písmena, čísla, tečku, pomlčku a podtržítko
                # Vše ostatní (včetně mezer) nahradí podtržítkem nebo odstraní
                filename_clean = re.sub(r'[^\w\.-]', '_', filename_ascii)
                
                # Odstranění vícenásobných podtržítek vzniklých čištěním
                filename_clean = re.sub(r'_+', '_', filename_clean)

                _LOGGER.info(f'Original filename: {filename} -> Cleaned: {filename_clean}')

                payload = part.get_payload(decode=True)
                if payload is None:
                    continue

                fullpath = os.path.join(storage_path, filename_clean)
                
                try:
                    with open(fullpath, 'wb') as f:
                        f.write(payload)
                    
                    attachments.append(fullpath)

                    # Logika pro Pandas (Excel -> CSV)
                    if filename_clean.lower().endswith(('.xls', '.xlsx')):
                        try:
                            # POZOR: Pandas musí číst ten nový "vyčištěný" název souboru
                            read_file = pd.read_excel(fullpath)
                            output_csv_path = os.path.join(storage_path, csv_filename)
                            read_file.to_csv(output_csv_path, index=None, header=True, sep=';')
                        except Exception as e:
                            _LOGGER.error(f"Pandas conversion failed: {e}")

                except Exception as e:
                    _LOGGER.error(f"Failed to save file {fullpath}: {e}")

        return attachments

    def update(self):
        _LOGGER.info('Update emails')
        """Read emails and publish state change."""
        email_message = self._email_reader.read_next()

        if email_message is None:
            self._message = None
            self._state_attributes = {}
            return

        if self.sender_allowed(email_message):
            message = EmailContentSensor.get_msg_subject(email_message)

            if self._value_template is not None:
                message = self.render_template(email_message)
            _LOGGER.info('Message')
            #_LOGGER.info(message)
            self._message = message
            self._state_attributes = {
                ATTR_FROM: EmailContentSensor.get_msg_sender(email_message),
                ATTR_SUBJECT: EmailContentSensor.get_msg_subject(email_message),
                ATTR_DATE: email_message["Date"],
                ATTR_BODY: EmailContentSensor.get_msg_text(email_message),
                ATTR_NUM_ATTACHMENTS: EmailContentSensor.get_num_msg_attachments(email_message),
                ATTR_ATTACHMENT_PATHS: EmailContentSensor.get_msg_attachments(email_message, self._storage_path, self._csv_filename),
            }

            _LOGGER.info('Smazat zprávu: ' + str(self._email_reader._last_id))
            self._email_reader.connection.uid("COPY", self._email_reader._last_uid, '"[Gmail]/Trash"')
            self._email_reader.connection.uid("STORE", self._email_reader._last_uid, "+FLAGS", r"(\Deleted)")
            #self._email_reader.connection.store(message, "+FLAGS", "\\Deleted")

