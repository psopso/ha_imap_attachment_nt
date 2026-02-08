"""Email attachment sensor support with Tariff Logic and Throttling."""
import logging
import os
import json
import re
import datetime
import imaplib
import email
import unicodedata
from email.header import decode_header, make_header

import voluptuous as vol
import pandas as pd

from homeassistant.components.sensor import PLATFORM_SCHEMA
from homeassistant.const import (
    CONF_NAME,
    CONF_PASSWORD,
    CONF_PORT,
    CONF_USERNAME,
    CONF_VALUE_TEMPLATE,
)
import homeassistant.helpers.config_validation as cv
from homeassistant.helpers.entity import Entity

_LOGGER = logging.getLogger(__name__)

# --- KONFIGURACE ---
CONF_SERVER = "server"
CONF_SENDERS = "senders"
CONF_FOLDER = "folder"
CONF_STORAGE_PATH = "storage_path"
CONF_EMAIL_INTERVAL = "email_interval_minutes"

DEFAULT_PORT = 993
DEFAULT_EMAIL_INTERVAL = 60 # Kontrola emailu jednou za hodinu

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
        vol.Optional(CONF_EMAIL_INTERVAL, default=DEFAULT_EMAIL_INTERVAL): cv.positive_int,
    }
)

def setup_platform(hass, config, add_entities, discovery_info=None):
    """Set up the Email sensor platform."""
    storage_path = config.get(CONF_STORAGE_PATH)
    if not os.path.exists(storage_path):
        os.makedirs(storage_path)

    processor = TariffProcessor(storage_path)
    
    reader = EmailReader(
        config.get(CONF_USERNAME),
        config.get(CONF_PASSWORD),
        config.get(CONF_SERVER),
        config.get(CONF_PORT),
        config.get(CONF_FOLDER),
        storage_path,
        processor 
    )

    sensor = EmailContentSensor(
        hass,
        reader,
        processor,
        config.get(CONF_NAME) or "Tariff Sensor",
        config.get(CONF_SENDERS),
        config.get(CONF_EMAIL_INTERVAL)
    )

    # OPRAVA: Odstraněna podmínka "if sensor.connected", která způsobovala chybu.
    # Sensor se přidá vždy, připojení se řeší až v metodě update().
    add_entities([sensor], True)


class TariffProcessor:
    """Třída pro zpracování Excelu a výpočet stavu tarifu (Lokální data)."""

    def __init__(self, storage_path):
        self.storage_path = storage_path
        self.json_file = os.path.join(storage_path, "tariff_data.json")
        
        self.days_map = {
            'Pondělí': 0, 'Úterý': 1, 'Středa': 2, 'Čtvrtek': 3, 
            'Pátek': 4, 'Sobota': 5, 'Neděle': 6
        }

    def process_excel(self, excel_path):
        """Načte Excel, vyfiltruje data a uloží je jako JSON."""
        try:
            _LOGGER.info(f"Processing Excel: {excel_path}")
            df = pd.read_excel(excel_path)
            
            if df.empty:
                return False

            # Filtrujeme řádky, kde první sloupec končí na '|1' (NT)
            mask = df.iloc[:, 0].astype(str).str.strip().str.endswith('|1')
            df_filtered = df[mask]

            tariff_schedule = {}

            for _, row in df_filtered.iterrows():
                day_name = row.iloc[2] # Sloupec Platnost
                if day_name not in self.days_map:
                    continue
                
                day_idx = self.days_map[day_name]
                intervals = []

                # Procházíme sloupce s časy po dvojicích
                for i in range(3, 13, 2):
                    try:
                        start = row.iloc[i]
                        end = row.iloc[i+1]
                        
                        if pd.notna(start) and pd.notna(end):
                            s_str = str(start).strip()
                            e_str = str(end).strip()
                            if s_str and e_str and s_str != 'nan':
                                intervals.append((s_str, e_str))
                    except IndexError:
                        break
                
                tariff_schedule[day_idx] = intervals

            with open(self.json_file, 'w') as f:
                json.dump(tariff_schedule, f, indent=4)
            
            _LOGGER.info("Tariff data saved successfully to JSON.")
            return True

        except Exception as e:
            _LOGGER.error(f"Failed to process Excel: {e}")
            return False

    def get_current_state(self):
        """Vypočítá, zda jsme v NT a jaký je aktuální/příští interval."""
        if not os.path.exists(self.json_file):
            return "Unknown", "Waiting for data (JSON missing)"

        try:
            with open(self.json_file, 'r') as f:
                data = json.load(f)
        except Exception:
            return "Unknown", "Error reading JSON data"

        now = datetime.datetime.now()
        current_day = now.weekday()
        
        def to_mins(t_str):
            if t_str == "24:00": return 1440
            try:
                h, m = map(int, t_str.split(':'))
                return h * 60 + m
            except:
                return 0

        current_mins = to_mins(now.strftime("%H:%M"))

        # --- LOGIKA SPOJOVÁNÍ INTERVALŮ ---
        timeline = []
        for day_offset in [-1, 0, 1]:
            target_day_idx = (current_day + day_offset) % 7
            key = str(target_day_idx)
            
            if key in data:
                offset_mins = day_offset * 1440
                for start_s, end_s in data[key]:
                    s_m = to_mins(start_s) + offset_mins
                    e_m = to_mins(end_s) + offset_mins
                    timeline.append({'start': s_m, 'end': e_m})

        timeline.sort(key=lambda x: x['start'])

        # Sloučení
        merged = []
        if timeline:
            curr = timeline[0]
            for i in range(1, len(timeline)):
                if timeline[i]['start'] <= curr['end']:
                    curr['end'] = max(curr['end'], timeline[i]['end'])
                else:
                    merged.append(curr)
                    curr = timeline[i]
            merged.append(curr)

        # Vyhodnocení
        active_interval = None
        next_interval = None

        for item in merged:
            if item['start'] <= current_mins < item['end']:
                active_interval = item
                break
            if item['start'] > current_mins:
                next_interval = item
                break

        def min_to_hm(m):
            d = 0
            if m >= 1440: d = 1; m -= 1440
            elif m < 0: d = -1; m += 1440
            t = f"{m//60:02d}:{m%60:02d}"
            if d == 1: return f"Zítra {t}"
            if d == -1: return f"Včera {t}"
            return t

        if active_interval:
            return "NT", f"{min_to_hm(active_interval['start'])} - {min_to_hm(active_interval['end'])}"
        elif next_interval:
            return "VT", f"Příští NT: {min_to_hm(next_interval['start'])} - {min_to_hm(next_interval['end'])}"
        
        return "VT", "Žádný další NT v dohledu"


class EmailReader:
    """Třída pro komunikaci s IMAP."""

    def __init__(self, user, password, server, port, folder, storage_path, processor):
        self._user = user
        self._password = password
        self._server = server
        self._port = port
        self._folder = folder
        self._storage_path = storage_path
        self._processor = processor
        self.connection = None
        self._last_uid = None

    def connect(self):
        try:
            self.connection = imaplib.IMAP4_SSL(self._server, self._port)
            self.connection.login(self._user, self._password)
            return True
        except Exception as e:
            _LOGGER.error(f"Failed to login: {e}")
            return False

    def check_for_emails(self):
        """Připojí se, stáhne emaily, uloží přílohy a odpojí se."""
        _LOGGER.info("Connecting to IMAP to check for new emails...")
        if not self.connect():
            return False

        try:
            self.connection.select(self._folder, readonly=False)
            search_crit = '(SUBJECT "Export" UNSEEN)'
            
            _, data = self.connection.uid("search", None, search_crit)
            uids = data[0].split()

            if not uids:
                _LOGGER.info("No new emails found.")
                self.connection.logout()
                return False

            self._last_uid = uids[-1]
            _, msg_data = self.connection.uid("fetch", self._last_uid, "(RFC822)")
            
            if msg_data and msg_data[0]:
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                
                self._save_attachments(email_message)

                try:
                    self.connection.uid("STORE", self._last_uid, "+FLAGS", r"(\Deleted)")
                    self.connection.expunge()
                    _LOGGER.info("Email processed and deleted.")
                except Exception as e:
                    _LOGGER.warning(f"Delete failed: {e}")

            self.connection.logout()
            return True

        except Exception as e:
            _LOGGER.error(f"Error during email check: {e}")
            try:
                self.connection.logout()
            except:
                pass
            return False

    def _save_attachments(self, email_message):
        """Interní metoda pro uložení příloh."""
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            filename = part.get_filename()
            if filename:
                try:
                    filename = str(make_header(decode_header(filename)))
                except: pass

                norm = unicodedata.normalize('NFKD', filename)
                ascii_name = norm.encode('ASCII', 'ignore').decode('ASCII')
                clean_name = re.sub(r'[^\w\.-]', '_', ascii_name)
                
                payload = part.get_payload(decode=True)
                if payload:
                    fullpath = os.path.join(self._storage_path, clean_name)
                    with open(fullpath, 'wb') as f:
                        f.write(payload)
                    
                    if clean_name.lower().endswith(('.xls', '.xlsx')):
                        self._processor.process_excel(fullpath)


class EmailContentSensor(Entity):
    """Hlavní sensor."""

    def __init__(self, hass, reader, processor, name, allowed_senders, email_interval):
        self.hass = hass
        self._reader = reader
        self._processor = processor
        self._name = name
        self._allowed_senders = [s.upper() for s in allowed_senders]
        
        # Konfigurace intervalu
        self._email_check_interval = datetime.timedelta(minutes=email_interval)
        # Nastavíme minulost, aby se stahovalo hned po startu
        self._last_email_check = datetime.datetime.min 
        
        self._state = "Unknown"
        self._attributes = {}

    @property
    def name(self):
        return self._name

    @property
    def state(self):
        return self._state

    @property
    def extra_state_attributes(self):
        return self._attributes

    def update(self):
        """Tato metoda se volá každých cca 30 sekund."""
        now = datetime.datetime.now()

        # 1. ČÁST: KONTROLA EMAILU (dle intervalu)
        if now - self._last_email_check > self._email_check_interval:
            _LOGGER.info(f"Time for scheduled email check (Interval: {self._email_check_interval})")
            # Zde probíhá skutečné připojení k serveru
            self._reader.check_for_emails()
            self._last_email_check = now
        
        # 2. ČÁST: PŘEPOČET TARIFU (vždy)
        state_val, interval_info = self._processor.get_current_state()
        
        self._state = state_val
        self._attributes = {
            "info": interval_info,
            "last_email_check": self._last_email_check.isoformat(),
            "next_email_check": (self._last_email_check + self._email_check_interval).isoformat()
        }
        