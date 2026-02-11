# Home Assistant – ČEZ Tariff Sensor (IMAP)

Tato integrace automatizuje sledování nízkého tarifu (HDO) na základě e-mailových exportů od společnosti ČEZ. Už nemusíte ručně hlídat časy v Excelu – senzor to udělá za vás.

## Hlavní funkce
- Automatický import: Sleduje e-mailovou schránku a stahuje přílohy (Excel) z e-mailů s předmětem "Export".
- Chytré zpracování: Filtruje data pro konkrétní tarify (povel zakončený |1) a ukládá je do lokálního JSON souboru.
- Offline odolnost: Po stažení dat funguje senzor nezávisle na internetu (čte z lokální paměti).
- Spojování intervalů: Pokud nízký tarif pokračuje i po půlnoci, senzor tyto intervaly spojí do jednoho logického bloku.
- Optimalizace výkonu: Kontrola e-mailu probíhá v dlouhých intervalech, zatímco stav senzoru se přepočítává bleskově lokálně.

## Instalace

1. Vytvořte složku /config/custom_components/imap_attachment_nt/.
2. Do této složky nahrajte soubory: sensor.py a manifest.json.
3. Restartujte Home Assistant.

## Konfigurace (configuration.yaml)

sensor:
  - platform: imap_attachment_nt
    name: "Elektřina Nízký Tarif"
    server: "imap.seznam.cz"
    port: 993
    username: "vas-email@seznam.cz"
    password: "vase-heslo-k-emailu"
    senders:
      - "info@cez.cz"
    storage_path: "/config/attachments"
    email_interval_minutes: 60

## Jak to funguje
- Stahování: V intervalu se senzor připojí k IMAPu a najde e-mail s "Export" v předmětu.
- Čištění: Názvy souborů jsou zbaveny diakritiky pro kompatibilitu s Linuxem.
- Analýza: Pomocí knihovny pandas se vyfiltruje nízký tarif (NT) pro každý den.
- Atributy: V atributu "info" uvidíte aktuální nebo příští časové okno.

## Závislosti
Integrace využívá knihovny pandas a openpyxl. Tyto knihovny si Home Assistant nainstaluje automaticky při prvním spuštění na základě manifest.json.