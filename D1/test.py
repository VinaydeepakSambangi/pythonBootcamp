from fpdf import FPDF
from docx import Document
from collections import Counter

a= [
  {
    "name": "2N Doorbell",
    "url": "https://driverstore.rticontrol.com/driver/2n",
    "categories": [
      "Access Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 14, 2022",
    "version": "1.0",
    "description": "2N Solo and IP Verso"
  },
  {
    "name": "DoorBird",
    "url": "https://driverstore.rticontrol.com/driver/doorbird-1",
    "categories": [
      "Access Control"
    ],
    "developer": "RTI",
    "updated": "June 5, 2026",
    "version": "3.1",
    "description": "Driver for DoorBird"
  },
  {
    "name": "DoorBird (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/doorbird",
    "categories": [
      "Access Control"
    ],
    "developer": "RTI",
    "updated": "July 26, 2018",
    "version": "1.13",
    "description": "DoorBird"
  },
  {
    "name": "Genie & Overhead Door",
    "url": "https://driverstore.rticontrol.com/driver/genie",
    "categories": [
      "Access Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2026",
    "version": "1.4",
    "description": "Driver for Genie & Overhead Door Garage Door Openers"
  },
  {
    "name": "ismartgate Garage",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-ismartgate",
    "categories": [
      "Access Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "240204.00",
    "description": "Local communication with an ismartgate Smart Opener"
  },
  {
    "name": "MyQ",
    "url": "https://driverstore.rticontrol.com/driver/myq",
    "categories": [
      "Access Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "260302.00",
    "description": "Driver for MYQ"
  },
  {
    "name": "RTI IVDS",
    "url": "https://driverstore.rticontrol.com/driver/rti-ivds",
    "categories": [
      "Access Control"
    ],
    "developer": "RTI",
    "updated": "May 23, 2025",
    "version": "1.0",
    "description": "Driver for RTI IVDS"
  },
  {
    "name": "Tailwind",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-tailwind",
    "categories": [
      "Access Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "231204.00",
    "description": "Communication with up to 3 Tailwind garage door openers"
  },
  {
    "name": "AudioControl BijouSeries",
    "url": "https://driverstore.rticontrol.com/driver/bijouseries",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Dec. 5, 2025",
    "version": "1.2",
    "description": "Driver for BijouSeries"
  },
  {
    "name": "AudioControl CM-Series (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol-cm-series",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "April 5, 2024",
    "version": "1.0",
    "description": "(Legacy) Driver for AudioControl CM-Series"
  },
  {
    "name": "AudioControl CMx-750 and Director Series",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol-cmx-750-and-director-series",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Nov. 13, 2026",
    "version": "2.01",
    "description": "Driver for AudioControl CMx-750 and Director Series"
  },
  {
    "name": "AudioControl DirectorSeries (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/audio-control-the-director",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "July 24, 2025",
    "version": "2.12",
    "description": "(Legacy) Driver for AudioControl DirectorSeries"
  },
  {
    "name": "AudioControl RS Series",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol-rs-series",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Nov. 13, 2026",
    "version": "1.00",
    "description": "Driver for AudioControl RS Series"
  },
  {
    "name": "Audio Technica ATDM-0604a",
    "url": "https://driverstore.rticontrol.com/driver/audio-technica-atdm-0604a",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Aug. 2, 2024",
    "version": "1.0",
    "description": "Driver for Audio Technica ATDM-0604a"
  },
  {
    "name": "Australian Monitor ZoneMix",
    "url": "https://driverstore.rticontrol.com/driver/am-zonemix",
    "categories": [
      "Amplifier"
    ],
    "developer": "Miantic AV",
    "updated": "Nov. 15, 2024",
    "version": "1.1",
    "description": "Driver for AM (Australian Monitor) ZoneMix 4/8"
  },
  {
    "name": "Blustream NPA100DA",
    "url": "https://driverstore.rticontrol.com/driver/blustream-npa100da",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Nov. 3, 2023",
    "version": "1.0",
    "description": "Driver for Blustream NPA100DA"
  },
  {
    "name": "Calypso CA-1000",
    "url": "https://driverstore.rticontrol.com/driver/calypso-ca-1000",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Oct. 30, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Ecler CA Series Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/ecler-ca-series-amplifier",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Aug. 14, 2020",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Klipsch C-Series Subwoofers",
    "url": "https://driverstore.rticontrol.com/driver/klipsch-c-series-subwoofers",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Aug. 3, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Rotel RC Series Preamplifier",
    "url": "https://driverstore.rticontrol.com/driver/rotel-rc-series-preamplifier",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Jan. 7, 2022",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "RTI AD DSP Matrix",
    "url": "https://driverstore.rticontrol.com/driver/rti-ad-dsp-matrix",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Dec. 5, 2025",
    "version": "1.1",
    "description": "AD-46-100, AD-810-100, AD-1616"
  },
  {
    "name": "RTI AMR-220 Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/rti-amr-220-amplifier",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI AMR-350 3 Input Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/rti-amr-350-3-input-amplifier",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI AMV-340 3 Input Constant Voltage Amplifier with Mic",
    "url": "https://driverstore.rticontrol.com/driver/rti-amv-340-3-input-constant-voltage-amplifier-with-mic",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI CP Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/rti-cp-amplifier",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Nov. 15, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI Multi-Z AMP",
    "url": "https://driverstore.rticontrol.com/driver/rti-multi-z-amp",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Aug. 28, 2025",
    "version": "1.20",
    "description": "Driver for RTI Multi-Z AMP"
  },
  {
    "name": "Stewart Audio AV-25",
    "url": "https://driverstore.rticontrol.com/driver/stewart-audio-av-25",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Aug. 16, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Trinnov Altitude",
    "url": "https://driverstore.rticontrol.com/driver/trinnov-altitude",
    "categories": [
      "Amplifier"
    ],
    "developer": "Janus Technology",
    "updated": "June 11, 2026",
    "version": "1.19",
    "description": "Driver for Trinnov Altitude"
  },
  {
    "name": "Velodyne Digital Drive",
    "url": "https://driverstore.rticontrol.com/driver/velodyne-digital-drive",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "March 26, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Velodyne SubContractor",
    "url": "https://driverstore.rticontrol.com/driver/velodyne-subcontractor",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "March 26, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "WiiM",
    "url": "https://driverstore.rticontrol.com/driver/wiim",
    "categories": [
      "Amplifier"
    ],
    "developer": "RTI",
    "updated": "Jan. 3, 2025",
    "version": "1.05",
    "description": "Driver for WiiM"
  },
  {
    "name": "ADA Suite 7.1",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-ada-suite-71",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "My Device",
    "updated": "June 7, 2014",
    "version": "1.0",
    "description": "Driver for ADA Suite 7.1"
  },
  {
    "name": "Anthem AVM 50v Processor",
    "url": "https://driverstore.rticontrol.com/driver/anthem-avm-50v-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Anthem MRX Receivers",
    "url": "https://driverstore.rticontrol.com/driver/anthem-mrx-receivers",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Anthem MRX x10 Series",
    "url": "https://driverstore.rticontrol.com/driver/anthem-mrx-x10-series",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 20, 2017",
    "version": "1.23",
    "description": ""
  },
  {
    "name": "Anthem MRXx40",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-anthem-mrxx40",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "1.5",
    "description": "This module provides control for the Anthem MRXx40 range of products"
  },
  {
    "name": "Arcam AVR",
    "url": "https://driverstore.rticontrol.com/driver/arcam-avr",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "Arcam AVR390/550/850/860",
    "url": "https://driverstore.rticontrol.com/driver/arcam-avr390-550-850-860",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "March 21, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Arcam AVR Series",
    "url": "https://driverstore.rticontrol.com/driver/arcam-avr-series",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Oct. 25, 2024",
    "version": "1.0",
    "description": "Driver for Arcam AVR Series"
  },
  {
    "name": "AudioControl AVR",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol-avr",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "AudioControl AVR-7/9",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol-avr-7-9",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Jan. 27, 2017",
    "version": "1.04",
    "description": ""
  },
  {
    "name": "Bose Lifetyle Home Entertainment System",
    "url": "https://driverstore.rticontrol.com/driver/bose-lifetyle-home-entertainment-system",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.32",
    "description": ""
  },
  {
    "name": "Cary Audio Cinema 11a",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-cary-audio-cinema-11a",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "My Device",
    "updated": "June 28, 2011",
    "version": "1.0",
    "description": "Driver for Cary Audio Cinema 11a"
  },
  {
    "name": "Classe Audio Surround Processor",
    "url": "https://driverstore.rticontrol.com/driver/classe-audio-surround-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Classe Sigma SSP Processor",
    "url": "https://driverstore.rticontrol.com/driver/classe-sigma-ssp-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Aug. 31, 2017",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "Datasat AP20",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-datasat-ap20",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "My Device",
    "updated": "June 7, 2013",
    "version": "1.0",
    "description": "Driver for Datasat AP20"
  },
  {
    "name": "Denon Marantz AVR",
    "url": "https://driverstore.rticontrol.com/driver/denon-marantz-avr",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Feb. 10, 2023",
    "version": "2.01",
    "description": ""
  },
  {
    "name": "DVDO iScan Duo",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-dvdo-iscan-duo",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "My Device",
    "updated": "July 26, 2013",
    "version": "1.0",
    "description": "Driver for DVDO iScan Duo"
  },
  {
    "name": "Dynaudio Focus",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-dynaudio",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "240604.00",
    "description": "Audio zone functionality and basic meta data feedback for Dynaudio"
  },
  {
    "name": "Integra/Onkyo/Pioneer Receiver  2017+",
    "url": "https://driverstore.rticontrol.com/driver/integra-onkyo-pioneer-receiver-2017",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Feb. 28, 2020",
    "version": "1.05",
    "description": ""
  },
  {
    "name": "Integra Receiver DTR (Deprecated)",
    "url": "https://driverstore.rticontrol.com/driver/integra-receiver-dtr-deprecated",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Feb. 21, 2017",
    "version": "2.34",
    "description": ""
  },
  {
    "name": "JBL Synthesis SDP-45",
    "url": "https://driverstore.rticontrol.com/driver/jbl-synthesis-sdp-45",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Dec. 19, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Klipsch Flexus",
    "url": "https://driverstore.rticontrol.com/driver/klipsch-flexus",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "April 3, 2026",
    "version": "1.0",
    "description": "Driver for Klipsch Flexus"
  },
  {
    "name": "Marantz Receiver SR",
    "url": "https://driverstore.rticontrol.com/driver/marantz-receiver-sr",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Jan. 6, 2012",
    "version": "2.0",
    "description": ""
  },
  {
    "name": "Meridian Surround Controller",
    "url": "https://driverstore.rticontrol.com/driver/meridian-surround-controller",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "NAD Receiver",
    "url": "https://driverstore.rticontrol.com/driver/nad-receiver",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Novastar H Series (Lite)",
    "url": "https://driverstore.rticontrol.com/driver/novastar-h-series-lite",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Aug. 1, 2025",
    "version": "1.0",
    "description": "Driver for Novastar H Series (Lite)"
  },
  {
    "name": "Parasound ZPre2",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-parasound-zpre2",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "My Device",
    "updated": "Oct. 8, 2011",
    "version": "1.0",
    "description": "Driver for Parasound ZPre2"
  },
  {
    "name": "Parsound Zpre2 Zone Preamplifier",
    "url": "https://driverstore.rticontrol.com/driver/parsound-zpre2-zone-preamplifier",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Pioneer Elite SC",
    "url": "https://driverstore.rticontrol.com/driver/pioneer-elite-sc",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2015",
    "version": "1.32",
    "description": ""
  },
  {
    "name": "Rotel RAP1580 Surround Amplified Processor",
    "url": "https://driverstore.rticontrol.com/driver/rotel-rap1580-surround-amplified-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "June 8, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Rotel RC-1570 Stereo Preamplifier",
    "url": "https://driverstore.rticontrol.com/driver/rotel-rc-1570-stereo-preamplifier",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Oct. 6, 2017",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Rotel Receiver",
    "url": "https://driverstore.rticontrol.com/driver/rotel-receiver",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "March 15, 2013",
    "version": "1.10",
    "description": ""
  },
  {
    "name": "Rotel RSP1582 Sound Processor",
    "url": "https://driverstore.rticontrol.com/driver/rotel-rsp1582-sound-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "June 8, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Samsung Soundbars",
    "url": "https://driverstore.rticontrol.com/driver/samsung-soundbars",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Sept. 27, 2024",
    "version": "1.0",
    "description": "Driver for Samsung Soundbars"
  },
  {
    "name": "Sony ES Receiver",
    "url": "https://driverstore.rticontrol.com/driver/sony-es-receiver",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Nov. 9, 2018",
    "version": "1.32",
    "description": ""
  },
  {
    "name": "Sony Sound Bars",
    "url": "https://driverstore.rticontrol.com/driver/sony-sound-bars",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2015",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Sony STR-ZA",
    "url": "https://driverstore.rticontrol.com/driver/sony-str-za",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Jan. 17, 2018",
    "version": "1.15",
    "description": ""
  },
  {
    "name": "Steinway-Lyngdorf Model P200 Surround Processor",
    "url": "https://driverstore.rticontrol.com/driver/steinway-lyngdorf-model-p200-surround-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "May 14, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Steinway-Lyngdorf Surround Processor",
    "url": "https://driverstore.rticontrol.com/driver/steinway-lyngdorf-surround-processor",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Storm Audio ISP",
    "url": "https://driverstore.rticontrol.com/driver/storm-audio-isp",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Nov. 4, 2022",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "Yamaha AVENTAGE Receiver",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-aventage-receiver",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "June 4, 2021",
    "version": "1.6",
    "description": ""
  },
  {
    "name": "Yamaha MusicCast",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-musiccast",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Oct. 16, 2026",
    "version": "3.1",
    "description": "Driver for Yamaha MusicCast"
  },
  {
    "name": "Yamaha MusicCast (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-musiccast-legacy",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "Aug. 23, 2024",
    "version": "2.22",
    "description": "Yamaha MusicCast"
  },
  {
    "name": "Yamaha Receiver RX-V",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-receiver-rx-v",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "June 3, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Yamaha Receiver RX-Z11",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-receiver-rx-z11",
    "categories": [
      "A/V Receiver"
    ],
    "developer": "RTI",
    "updated": "June 3, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Advantage Air",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-advantage-air",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.0",
    "description": "Control your Advantage Air System from RTI including support for systems with independent zone control."
  },
  {
    "name": "Advantage Air Gen III",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-advantage-air-gen-iii",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "My Device",
    "updated": "July 26, 2013",
    "version": "1.2",
    "description": "Driver for Advantage Air Gen III"
  },
  {
    "name": "Aprilaire",
    "url": "https://driverstore.rticontrol.com/driver/aprilaire",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Feb. 9, 2024",
    "version": "2.05",
    "description": ""
  },
  {
    "name": "Aprilaire WiFi Thermostats",
    "url": "https://driverstore.rticontrol.com/driver/aprilaire-wifi-thermostats",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Feb. 11, 2021",
    "version": "1.06",
    "description": ""
  },
  {
    "name": "CoolMaster",
    "url": "https://driverstore.rticontrol.com/driver/coolmaster",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "May 3, 2013",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "CoolMasterNet",
    "url": "https://driverstore.rticontrol.com/driver/coolmasternet",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "March 18, 2022",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Dooya",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-dooya",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.0",
    "description": "Driver for Dooya None"
  },
  {
    "name": "Ecobee Thermostats",
    "url": "https://driverstore.rticontrol.com/driver/ecobee-thermostats",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 14, 2026",
    "version": "1.08",
    "description": ""
  },
  {
    "name": "Hayward OmniLogic",
    "url": "https://driverstore.rticontrol.com/driver/hayward-omnilogic",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 19, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Heatmiser Neo",
    "url": "https://driverstore.rticontrol.com/driver/heatmiser_neo",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Janus Technology",
    "updated": "July 10, 2026",
    "version": "1.01",
    "description": "Driver for Heatmiser_Neo"
  },
  {
    "name": "Heatmiser NeoHub Pro",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-heatmiserneo",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "260302.00",
    "description": "Heatmiser NeoHub full thermostat heating control"
  },
  {
    "name": "Jandy AquaLink RS",
    "url": "https://driverstore.rticontrol.com/driver/jandy-aqualink-rs",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 15, 2010",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Jandy iAquaLink",
    "url": "https://driverstore.rticontrol.com/driver/jandy-iaqualink",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 14, 2023",
    "version": "1.23",
    "description": ""
  },
  {
    "name": "Netatmo Weather Station",
    "url": "https://driverstore.rticontrol.com/driver/netatmo-weather-station",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "May 15, 2020",
    "version": "1.33",
    "description": ""
  },
  {
    "name": "Network Thermostat",
    "url": "https://driverstore.rticontrol.com/driver/network-thermostat",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "March 28, 2025",
    "version": "1.0",
    "description": "Driver for Network Thermostat"
  },
  {
    "name": "Pentair IntelliTouch",
    "url": "https://driverstore.rticontrol.com/driver/pentair-intellitouch",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "April 12, 2013",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Polyaire Airtouch5",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-polyaire-airtouch5",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Chowmain Software",
    "updated": "Nov. 21, 2025",
    "version": "1.0",
    "description": "The Polyaire Airtouch5 driver enables integration of the RTI control system with the Airtouch5 zoning system."
  },
  {
    "name": "Radio Thermostat Company",
    "url": "https://driverstore.rticontrol.com/driver/radio-thermostat-company",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Nov. 21, 2012",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "RTI Works with Nest",
    "url": "https://driverstore.rticontrol.com/driver/rti-works-with-nest",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "March 9, 2018",
    "version": "1.31",
    "description": ""
  },
  {
    "name": "Sensibo",
    "url": "https://driverstore.rticontrol.com/driver/sensibo",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "4.0",
    "description": "Easy A/C control from many vendors for RTI"
  },
  {
    "name": "Tekmar",
    "url": "https://driverstore.rticontrol.com/driver/tekmar",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "April 10, 2026",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Venstar ColorTouch",
    "url": "https://driverstore.rticontrol.com/driver/venstar-colortouch",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "July 18, 2017",
    "version": "1.23",
    "description": ""
  },
  {
    "name": "Weather",
    "url": "https://driverstore.rticontrol.com/driver/weather",
    "categories": [
      "Climate and Pool Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 4, 2024",
    "version": "3.04",
    "description": ""
  },
  {
    "name": "Denon Blu-Ray Player DBP-2012UD",
    "url": "https://driverstore.rticontrol.com/driver/denon-blu-ray-player-dbp-2012ud",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "Jan. 2, 2012",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Integra Blu-Ray/DVD",
    "url": "https://driverstore.rticontrol.com/driver/integra-blu-ray-dvd",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "Sept. 17, 2012",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Magnetar Blu-ray Players",
    "url": "https://driverstore.rticontrol.com/driver/magnetar-bluray-players",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "May 24, 2024",
    "version": "1.0",
    "description": "Driver for Magnetar Bluray Players"
  },
  {
    "name": "OPPO Bluray",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-oppo-bluray",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "My Device",
    "updated": "June 6, 2013",
    "version": "3.2",
    "description": "Driver for OPPO Bluray"
  },
  {
    "name": "Pioneer BDP-LX52",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-pioneer-bdp-lx52",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "My Device",
    "updated": "May 25, 2012",
    "version": "1.1",
    "description": "Driver for Pioneer BDP-LX52"
  },
  {
    "name": "Pioneer Elite BDP",
    "url": "https://driverstore.rticontrol.com/driver/pioneer-elite-bdp",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "July 20, 2017",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Pioneer UDP-LX Blu-ray Player",
    "url": "https://driverstore.rticontrol.com/driver/pioneer-udp-lx-blu-ray-player",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "March 8, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Samsung IP UHD Bluray",
    "url": "https://driverstore.rticontrol.com/driver/samsung-ip-uhd-bluray",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "July 23, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Sharp BD-AMS20U Bluray",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-sharp-bd-ams20u-bluray",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "My Device",
    "updated": "Oct. 6, 2012",
    "version": "1.0",
    "description": "Driver for Sharp BD-AMS20U Bluray"
  },
  {
    "name": "Sony BDP-CX7000ES",
    "url": "https://driverstore.rticontrol.com/driver/sony-bdp-cx7000es",
    "categories": [
      "Disc Player / Changer"
    ],
    "developer": "RTI",
    "updated": "March 18, 2011",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Aquavision Nexus Televisions",
    "url": "https://driverstore.rticontrol.com/driver/aquavision-nexus-televisions",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "July 7, 2017",
    "version": "1.0",
    "description": "Aquavision Nexus Televisions(Legacy)"
  },
  {
    "name": "Aquavision Television",
    "url": "https://driverstore.rticontrol.com/driver/aquavision-television",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Aug. 11, 2023",
    "version": "2.0",
    "description": "Driver for Aquavision Television"
  },
  {
    "name": "Dynamic Camera Feed",
    "url": "https://driverstore.rticontrol.com/driver/dynamic-camera-feed",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "May 26, 2026",
    "version": "1.0",
    "description": "Driver for Dynamic Camera Feed"
  },
  {
    "name": "Everpass Control API",
    "url": "https://driverstore.rticontrol.com/driver/control-concepts-everpass-control-api",
    "categories": [
      "Display"
    ],
    "developer": "Control Concepts, Inc.",
    "updated": "Aug. 18, 2026",
    "version": "1.0",
    "description": "Driver for Everpass Control API"
  },
  {
    "name": "Hisense Commercial Display",
    "url": "https://driverstore.rticontrol.com/driver/hisense-commercial-display",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Oct. 23, 2026",
    "version": "1.0",
    "description": "Driver for Hisense Commercial Display"
  },
  {
    "name": "Hisense Smart TV",
    "url": "https://driverstore.rticontrol.com/driver/hisense-smart-tv",
    "categories": [
      "Display"
    ],
    "developer": "Hisense USA Corporation",
    "updated": "Feb. 6, 2026",
    "version": "1.3",
    "description": "Driver for Hisense Smart TV"
  },
  {
    "name": "Hisense V-HOME Smart TV",
    "url": "https://driverstore.rticontrol.com/driver/v-home-smart-tv",
    "categories": [
      "Display"
    ],
    "developer": "Hisense USA Corporation",
    "updated": "July 10, 2026",
    "version": "1.0",
    "description": "Driver for V-HOME Smart TV"
  },
  {
    "name": "Just Video Walls H Series",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-justvideowalls-h-series",
    "categories": [
      "Display"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "250701.02",
    "description": "Control and one-way feedback for JustVideoWalls H-Series"
  },
  {
    "name": "Just Video Walls M Series",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-justvideowalls",
    "categories": [
      "Display"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "241103.01",
    "description": "Control and feedback of Just Video Walls M Series via IP"
  },
  {
    "name": "LG IP",
    "url": "https://driverstore.rticontrol.com/driver/lg-ip",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Aug. 28, 2026",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "LG IP",
    "url": "https://driverstore.rticontrol.com/driver/lg-ip-1",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2026",
    "version": "1.0.0 ICP",
    "description": "Driver for LG IP"
  },
  {
    "name": "LG Television",
    "url": "https://driverstore.rticontrol.com/driver/lg-television",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "March 30, 2018",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Loewe sl8 Series",
    "url": "https://driverstore.rticontrol.com/driver/loewe_sl8-series",
    "categories": [
      "Display"
    ],
    "developer": "Janus Technology",
    "updated": "July 24, 2026",
    "version": "1.01",
    "description": "Driver for Loewe_sl8-series"
  },
  {
    "name": "MAXHUB CMA and UW Series",
    "url": "https://driverstore.rticontrol.com/driver/maxhub-cma-and-uw-series",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "March 13, 2026",
    "version": "1.0",
    "description": "Driver for MAXHUB CMA and UW Series"
  },
  {
    "name": "NEC_P402 Display",
    "url": "https://driverstore.rticontrol.com/driver/nec_p402-display",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "May 17, 2019",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Panasonic IP TV",
    "url": "https://driverstore.rticontrol.com/driver/panasonic-ip-tv",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "July 24, 2020",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Panasonic Pro Plasma Displays",
    "url": "https://driverstore.rticontrol.com/driver/panasonic-pro-plasma-displays",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "March 26, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Panasonic Viera",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-panasonic-viera",
    "categories": [
      "Display"
    ],
    "developer": "My Device",
    "updated": "Sept. 8, 2015",
    "version": "1.1",
    "description": "Driver for Panasonic Viera"
  },
  {
    "name": "Philips BDL TV",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-philips-bdl-tv",
    "categories": [
      "Display"
    ],
    "developer": "Coupe Automation",
    "updated": "April 1, 2026",
    "version": "1.01",
    "description": "Driver for Philips BDL TV"
  },
  {
    "name": "Samsung Commercial Displays",
    "url": "https://driverstore.rticontrol.com/driver/samsung-commercial-displays",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Sept. 1, 2023",
    "version": "1.0",
    "description": "Driver for Samsung Commercial Displays"
  },
  {
    "name": "Samsung Ex-Link",
    "url": "https://driverstore.rticontrol.com/driver/samsung-ex-link",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Feb. 8, 2024",
    "version": "1.16",
    "description": ""
  },
  {
    "name": "Samsung IP",
    "url": "https://driverstore.rticontrol.com/driver/samsung-ip",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2026",
    "version": "1.0.0 ICP",
    "description": "Driver for Samsung IP"
  },
  {
    "name": "Samsung IP Television",
    "url": "https://driverstore.rticontrol.com/driver/samsung-ip-television",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "May 24, 2024",
    "version": "1.09",
    "description": ""
  },
  {
    "name": "Samsung Tizen OS TV with SmartThings",
    "url": "https://driverstore.rticontrol.com/driver/samsung-tizen-os-tv-with-smartthings",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2026",
    "version": "2.1",
    "description": "Driver for Samsung Tizen OS TV with SmartThings"
  },
  {
    "name": "Sharp Elite",
    "url": "https://driverstore.rticontrol.com/driver/sharp-elite",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "March 16, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "SkyQ Lite",
    "url": "https://driverstore.rticontrol.com/driver/skyq-lite",
    "categories": [
      "Display"
    ],
    "developer": "Clever AV Solutions",
    "updated": "July 12, 2024",
    "version": "3.0",
    "description": "Driver for SkyQ Lite"
  },
  {
    "name": "SMART Board 6000S Series",
    "url": "https://driverstore.rticontrol.com/driver/smart-board-6000s-series",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "May 12, 2023",
    "version": "1.0",
    "description": "Driver for SMART Board 6000S Series"
  },
  {
    "name": "Sony Bravia IP",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-ip-2",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Oct. 19, 2026",
    "version": "2.0.0 ICP",
    "description": "Driver for Sony Bravia IP"
  },
  {
    "name": "Sony Bravia IP",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-ip-1",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2026",
    "version": "2.12",
    "description": "Driver for Sony Bravia IP"
  },
  {
    "name": "Sony Bravia IP (Non-PSK Models)",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-ip",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Aug. 19, 2021",
    "version": "1.35",
    "description": "Sony Bravia IP"
  },
  {
    "name": "Sony Bravia Simple IP",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-simple-ip",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Dec. 27, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sony Bravia Simple IP",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-simple-ip-1",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2026",
    "version": "1.0.0 ICP",
    "description": "Driver for Sony Bravia Simple IP"
  },
  {
    "name": "Sony BRAVIA with Google TV",
    "url": "https://driverstore.rticontrol.com/driver/sony-bravia-with-google-tv",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "May 26, 2022",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Sony XBR Series LCD",
    "url": "https://driverstore.rticontrol.com/driver/sony-xbr-series-lcd",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Oct. 12, 2015",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Sunbrite Solis IP",
    "url": "https://driverstore.rticontrol.com/driver/sunbrite-solis-ip",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "July 3, 2025",
    "version": "1.0",
    "description": "Driver for Sunbrite Solis IP"
  },
  {
    "name": "SunBrite TV",
    "url": "https://driverstore.rticontrol.com/driver/sunbrite-tv",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "July 28, 2023",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "SunBrite Veranda 3",
    "url": "https://driverstore.rticontrol.com/driver/sunbrite-veranda-3",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Nov. 21, 2025",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Tizen OS for Samsung TVs",
    "url": "https://driverstore.rticontrol.com/driver/tizen-os-for-samsung-tvs",
    "categories": [
      "Display"
    ],
    "developer": "RTI",
    "updated": "Feb. 9, 2024",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "Allen and Heath AHM",
    "url": "https://driverstore.rticontrol.com/driver/allen-and-heath-ahm",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Aug. 28, 2026",
    "version": "1.21",
    "description": "Driver for Allen and Heath AHM"
  },
  {
    "name": "Atlas Atmosphere",
    "url": "https://driverstore.rticontrol.com/driver/atlas-atmosphere-1",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2026",
    "version": "2.23",
    "description": "Driver for Atlas Atmosphere"
  },
  {
    "name": "AtlasIED Atmosphere (Deprecated)",
    "url": "https://driverstore.rticontrol.com/driver/atlas-atmosphere",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Oct. 7, 2022",
    "version": "1.0",
    "description": "AtlasIED AZM4 and AZM8"
  },
  {
    "name": "Beale Street Audio BAV",
    "url": "https://driverstore.rticontrol.com/driver/beale-street-audio-bav",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Sept. 16, 2020",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Biamp Nexia",
    "url": "https://driverstore.rticontrol.com/driver/biamp-nexia",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Feb. 19, 2013",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Biamp Tesira",
    "url": "https://driverstore.rticontrol.com/driver/biamp-tesira",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "March 13, 2026",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Bose ControlSpace",
    "url": "https://driverstore.rticontrol.com/driver/bose-controlspace",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "June 20, 2025",
    "version": "2.21",
    "description": ""
  },
  {
    "name": "BSS Audio London BLU v1",
    "url": "https://driverstore.rticontrol.com/driver/bss-audio-london-blu",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Aug. 23, 2018",
    "version": "1.21",
    "description": "BSS Audio London BLU"
  },
  {
    "name": "BSS Audio London BLU v2",
    "url": "https://driverstore.rticontrol.com/driver/bss-audio-london-blu-2",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "March 10, 2023",
    "version": "2.01",
    "description": "BSS Audio London BLU"
  },
  {
    "name": "ClearOne Converge",
    "url": "https://driverstore.rticontrol.com/driver/clearone-converge",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Sept. 8, 2014",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "ClearOne CONVERGE Pro 2",
    "url": "https://driverstore.rticontrol.com/driver/clearone-converge-pro-2",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "July 9, 2019",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Cloud DCM-1/e",
    "url": "https://driverstore.rticontrol.com/driver/cloud-dcm-1-e",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "June 6, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Digital View DRR-2T DAB",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-digital-view-drr-2t-dab",
    "categories": [
      "DSP"
    ],
    "developer": "My Device",
    "updated": "June 21, 2012",
    "version": "1.0",
    "description": "Driver for Digital View DRR-2T DAB"
  },
  {
    "name": "Ecler DAM614",
    "url": "https://driverstore.rticontrol.com/driver/ecler-dam614",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Aug. 23, 2019",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Ecler eMIMO1616",
    "url": "https://driverstore.rticontrol.com/driver/ecler-emimo1616",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Nov. 3, 2021",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Ecler MIM01212SG",
    "url": "https://driverstore.rticontrol.com/driver/ecler-mim01212sg",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "June 3, 2022",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "Ecler MIMO4040DN",
    "url": "https://driverstore.rticontrol.com/driver/ecler-mimo4040dn",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Nov. 3, 2021",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Ecler MIMO88SG",
    "url": "https://driverstore.rticontrol.com/driver/ecler-mimo88sg",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Nov. 3, 2021",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Ecler VEO-XRI1C",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-ecler-veo-xxi1c",
    "categories": [
      "DSP"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "1.00",
    "description": "Channel selection on the VEO-XRI1C receiver via IP"
  },
  {
    "name": "Ecler VEO-XRI2L",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-ecler-veo-xxi2l",
    "categories": [
      "DSP"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "1.00",
    "description": "Provides control over the VEO-XRI2L receiver via IP"
  },
  {
    "name": "LEA Connect 168D",
    "url": "https://driverstore.rticontrol.com/driver/lea-connect-168d",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "July 15, 2022",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "Polycom SoundStructure",
    "url": "https://driverstore.rticontrol.com/driver/polycom-soundstructure",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Nov. 5, 2018",
    "version": "2.0",
    "description": ""
  },
  {
    "name": "PTN PA2 / ISLink DPA-22",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-ptn-pa2-islink-dpa-22",
    "categories": [
      "DSP"
    ],
    "developer": "My Device",
    "updated": "April 13, 2011",
    "version": "1.2",
    "description": "Driver for PTN PA2 / ISLink DPA-22"
  },
  {
    "name": "QSC MP-M",
    "url": "https://driverstore.rticontrol.com/driver/qsc-mp-m",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Dec. 13, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "QSC Q-SYS",
    "url": "https://driverstore.rticontrol.com/driver/qsc-q-sys",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Nov. 4, 2022",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Shure P300",
    "url": "https://driverstore.rticontrol.com/driver/shure-p300",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Sept. 25, 2026",
    "version": "1.0",
    "description": "Driver for Shure P300"
  },
  {
    "name": "Symetrix",
    "url": "https://driverstore.rticontrol.com/driver/symetrix",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "Aug. 26, 2016",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "TOA M-9000M2 Pre-Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/toa-m-9000m2-pre-amplifier",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "May 27, 2016",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Yamaha MRX",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-mrx",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "July 26, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Yamaha MTX",
    "url": "https://driverstore.rticontrol.com/driver/yamaha-mtx",
    "categories": [
      "DSP"
    ],
    "developer": "RTI",
    "updated": "June 15, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Yamaha YSP-4300",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-yamaha-ysp-4300",
    "categories": [
      "DSP"
    ],
    "developer": "My Device",
    "updated": "Oct. 3, 2014",
    "version": "1.0",
    "description": "Driver for Yamaha YSP-4300"
  },
  {
    "name": "Digital View Set Top",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-digital-view-set-top",
    "categories": [
      "DVR"
    ],
    "developer": "My Device",
    "updated": "July 11, 2010",
    "version": "1.0",
    "description": "Driver for Digital View Set Top"
  },
  {
    "name": "Humax",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-humax",
    "categories": [
      "DVR"
    ],
    "developer": "My Device",
    "updated": "Jan. 19, 2016",
    "version": "1.0",
    "description": "Driver for Humax"
  },
  {
    "name": "Reavon UBR Series",
    "url": "https://driverstore.rticontrol.com/driver/reavon-ubr-series",
    "categories": [
      "DVR"
    ],
    "developer": "RTI",
    "updated": "March 24, 2023",
    "version": "1.0",
    "description": "Driver for Reavon UBR Series"
  },
  {
    "name": "TiVo",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-tivo",
    "categories": [
      "DVR"
    ],
    "developer": "My Device",
    "updated": "Feb. 12, 2012",
    "version": "1.4",
    "description": "Driver for TiVo"
  },
  {
    "name": "TiVo Series 3 / HD",
    "url": "https://driverstore.rticontrol.com/driver/tivo-series-3-hd",
    "categories": [
      "DVR"
    ],
    "developer": "RTI",
    "updated": "July 25, 2017",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Xbox One Power",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-xbox-one-power",
    "categories": [
      "DVR"
    ],
    "developer": "My Device",
    "updated": "June 4, 2015",
    "version": "1.0",
    "description": "Driver for Xbox One Power"
  },
  {
    "name": "Aquameta",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-aquameta",
    "categories": [
      "Irrigation / Sprinklers"
    ],
    "developer": "My Device",
    "updated": "June 21, 2012",
    "version": "1.0",
    "description": "Driver for Aquameta"
  },
  {
    "name": "Irrigation Caddy",
    "url": "https://driverstore.rticontrol.com/driver/irrigation-caddy",
    "categories": [
      "Irrigation / Sprinklers"
    ],
    "developer": "RTI",
    "updated": "Oct. 31, 2014",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Rachio Iro Irrigation Controller",
    "url": "https://driverstore.rticontrol.com/driver/rachio-iro-irrigation-controller",
    "categories": [
      "Irrigation / Sprinklers"
    ],
    "developer": "RTI",
    "updated": "May 17, 2018",
    "version": "1.06",
    "description": ""
  },
  {
    "name": "Rain8net Irrigation Controller",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-rain8net-irrigation-controller",
    "categories": [
      "Irrigation / Sprinklers"
    ],
    "developer": "My Device",
    "updated": "Aug. 14, 2012",
    "version": "1.4",
    "description": "Driver for Rain8net Irrigation Controller"
  },
  {
    "name": "3 Day Blinds",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-3dayblinds",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "240508.00",
    "description": "Full control and feedback for up to 32 3DayBlinds devices through the 3 Day Blinds USB Hub."
  },
  {
    "name": "Black Nova Keypads",
    "url": "https://driverstore.rticontrol.com/driver/black-nova-keypads",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "June 30, 2023",
    "version": "1.0",
    "description": "Driver for Black Nova Keypads"
  },
  {
    "name": "CentraLite JetStream",
    "url": "https://driverstore.rticontrol.com/driver/centralite-jetstream",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "May 14, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Clipsal C-Bus",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-clipsal-c-bus",
    "categories": [
      "Lighting Control"
    ],
    "developer": "My Device",
    "updated": "Feb. 21, 2014",
    "version": "1.20",
    "description": "Driver for Clipsal C-Bus"
  },
  {
    "name": "Clipsal C-Bus",
    "url": "https://driverstore.rticontrol.com/driver/clipsal-c-bus",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 8, 2010",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Converging Systems Intelligent Lighting Controller",
    "url": "https://driverstore.rticontrol.com/driver/converging-systems-intelligent-lighting-controller",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "March 22, 2018",
    "version": "2.0",
    "description": "Converging Systems & PureEdge Intelligent Lighting Controller"
  },
  {
    "name": "Cooper Wiring Devices Aspire RF",
    "url": "https://driverstore.rticontrol.com/driver/cooper-wiring-devices-aspire-rf",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Jan. 22, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Current Products Corp Curtains",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-cpc-cloud",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "2.03",
    "description": "Two-way control of Current Products Corp curtains trac"
  },
  {
    "name": "Dynalite",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-dynalite",
    "categories": [
      "Lighting Control"
    ],
    "developer": "My Device",
    "updated": "Dec. 20, 2010",
    "version": "1.29",
    "description": "Driver for Dynalite"
  },
  {
    "name": "Ezlo Smart Home Hubs",
    "url": "https://driverstore.rticontrol.com/driver/ezlo-smart-home-hubs",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "April 25, 2025",
    "version": "2.07",
    "description": "Driver for Ezlo Smart Home Hubs"
  },
  {
    "name": "Fibaro HC3",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-fibarohc3",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "240212.00",
    "description": "Communication with Fibaro Home Center 3"
  },
  {
    "name": "FXLuminaire Luxor",
    "url": "https://driverstore.rticontrol.com/driver/fxluminaire-luxor",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 17, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Interactive Technologies CueServer 2",
    "url": "https://driverstore.rticontrol.com/driver/interactive-technologies-cueserver-2",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "July 20, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Legrand Shading Systems QzHub",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-qzhub-driver",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "230502.00",
    "description": "Integration and control of QMotion QzHub shades system"
  },
  {
    "name": "Legrand Shading Systems QzHub3",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-qmotion-qzhub3",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "2.03",
    "description": "Integration and two-way control of a QzHub3 shades system"
  },
  {
    "name": "Leviton ViziaRF",
    "url": "https://driverstore.rticontrol.com/driver/leviton-viziarf",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 21, 2012",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "LightwaveRF Lighting",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-lightwaverf",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "240214.00",
    "description": "Full control and feedback of up to 96 lighting/socket devices"
  },
  {
    "name": "LiteTouch 5000LC",
    "url": "https://driverstore.rticontrol.com/driver/litetouch-5000lc",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 15, 2010",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Lutron Caseta / RA2 Select",
    "url": "https://driverstore.rticontrol.com/driver/lutron-caseta-ra2-select",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "April 2, 2026",
    "version": "2.31",
    "description": ""
  },
  {
    "name": "Lutron Grafik Eye",
    "url": "https://driverstore.rticontrol.com/driver/lutron-grafik-eye",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Lutron Homeworks",
    "url": "https://driverstore.rticontrol.com/driver/lutron-homeworks",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 7, 2012",
    "version": "2.02",
    "description": "Lutron Homeworks Illumination"
  },
  {
    "name": "Lutron Homeworks RF",
    "url": "https://driverstore.rticontrol.com/driver/lutron-homeworks-rf",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "April 29, 2011",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Lutron LEAP",
    "url": "https://driverstore.rticontrol.com/driver/lutron-leap-1",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2026",
    "version": "2.51",
    "description": "Lutron RadioRA3, Homeworks QSX, and Athena"
  },
  {
    "name": "Lutron LEAP (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/lutron-leap",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "June 24, 2022",
    "version": "1.17",
    "description": "Lutron RadioRA3 and Homeworks QSX (deprecated, new installs should use driver v2)"
  },
  {
    "name": "Lutron myRoom",
    "url": "https://driverstore.rticontrol.com/driver/lutron-myroom",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2018",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Lutron RadioRA",
    "url": "https://driverstore.rticontrol.com/driver/lutron-radiora",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 30, 2009",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Lutron RadioRA 2 / Homeworks QS",
    "url": "https://driverstore.rticontrol.com/driver/lutron-radiora-2-homeworks-qs",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "March 31, 2023",
    "version": "2.26",
    "description": ""
  },
  {
    "name": "Lutron Vive",
    "url": "https://driverstore.rticontrol.com/driver/lutron-vive",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 8, 2023",
    "version": "1.11",
    "description": "Driver for Lutron Vive"
  },
  {
    "name": "MotionBlinds",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-motionblinds",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "240508.00",
    "description": "Full control and feedback of up to 32 MotionBlinds devices"
  },
  {
    "name": "Nice DMBM",
    "url": "https://driverstore.rticontrol.com/driver/nice-dmbm",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 6, 2024",
    "version": "1.1",
    "description": "Driver for Nice DMBM Module"
  },
  {
    "name": "NuLED SpectraDrive",
    "url": "https://driverstore.rticontrol.com/driver/nuled-spectradrive",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Jan. 11, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "On-Q RFLC Lighting",
    "url": "https://driverstore.rticontrol.com/driver/on-q-rflc-lighting",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Jan. 1, 2010",
    "version": "",
    "description": ""
  },
  {
    "name": "Orro Lighting",
    "url": "https://driverstore.rticontrol.com/driver/orro-lighting",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 16, 2022",
    "version": "1.11",
    "description": "Orro One Light Switch"
  },
  {
    "name": "Philips Hue",
    "url": "https://driverstore.rticontrol.com/driver/philips-hue",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 30, 2026",
    "version": "3.11",
    "description": "Driver for Philips Hue"
  },
  {
    "name": "Philips Hue Lighting (Deprecated)",
    "url": "https://driverstore.rticontrol.com/driver/philips-hue-lighting",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.21",
    "description": "Philips Hue Lighting"
  },
  {
    "name": "Philips Hue Lighting System (v2.X)",
    "url": "https://driverstore.rticontrol.com/driver/hue",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "May 24, 2024",
    "version": "2.13",
    "description": "Driver for Hue"
  },
  {
    "name": "PowerShades",
    "url": "https://driverstore.rticontrol.com/driver/powershades",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 10, 2025",
    "version": "1.10",
    "description": ""
  },
  {
    "name": "QMotion QConnect",
    "url": "https://driverstore.rticontrol.com/driver/qmotion-qconnect",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 5, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Qmotion QIS",
    "url": "https://driverstore.rticontrol.com/driver/qmotion-qis",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 15, 2016",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Qmotion Qsync",
    "url": "https://driverstore.rticontrol.com/driver/qmotion-qsync",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "June 23, 2016",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Qmotion Qube",
    "url": "https://driverstore.rticontrol.com/driver/qmotion-qube",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 7, 2016",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "QMotion Shades",
    "url": "https://driverstore.rticontrol.com/driver/qmotion-shades",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Jan. 1, 2010",
    "version": "",
    "description": ""
  },
  {
    "name": "Rako Lighting",
    "url": "https://driverstore.rticontrol.com/driver/rako-lighting",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "June 21, 2024",
    "version": "1.01",
    "description": "Driver for Rako Lighting"
  },
  {
    "name": "Rako Wireless",
    "url": "https://driverstore.rticontrol.com/driver/rako-wireless",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 28, 2011",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Rollease Pulse 2",
    "url": "https://driverstore.rticontrol.com/driver/rollease-pulse-2",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "July 10, 2020",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Rollease Pulse Hub",
    "url": "https://driverstore.rticontrol.com/driver/rollease-pulse-hub",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 17, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI KNX IP Driver",
    "url": "https://driverstore.rticontrol.com/driver/rti-knx-ip-driver",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 23, 2024",
    "version": "3.44",
    "description": "KNX IP Gateways"
  },
  {
    "name": "Shelly",
    "url": "https://driverstore.rticontrol.com/driver/shelly",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "June 12, 2026",
    "version": "1.31",
    "description": "Driver for Shelly"
  },
  {
    "name": "Silent Gliss Shades",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-slient-gliss",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "2.00",
    "description": "Two-way control of Silent Gliss shades via IP"
  },
  {
    "name": "Somfy Connect UAI Plus",
    "url": "https://driverstore.rticontrol.com/driver/somfy-connect-uai-plus",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "July 26, 2018",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Somfy ILT",
    "url": "https://driverstore.rticontrol.com/driver/somfy-ilt",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "March 18, 2011",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Somfy myLink",
    "url": "https://driverstore.rticontrol.com/driver/somfy-mylink",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 19, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Somfy POE Motor",
    "url": "https://driverstore.rticontrol.com/driver/somfy-poe-motor",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Control Concepts, Inc.",
    "updated": "April 22, 2025",
    "version": "1.2",
    "description": "Driver for Somfy POE Motor"
  },
  {
    "name": "Somfy TaHoma (EU)",
    "url": "https://driverstore.rticontrol.com/driver/somfy_tahoma",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Janus Technology",
    "updated": "Oct. 23, 2026",
    "version": "1.02",
    "description": "Driver for Somfy TaHoma EU"
  },
  {
    "name": "Somfy TaHoma North America",
    "url": "https://driverstore.rticontrol.com/driver/somfy-tahoma-north-america",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Control Concepts, Inc.",
    "updated": "Sept. 24, 2021",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "SwitchBot",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-switchbot",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "260302.00",
    "description": "SwitchBot communication via SwitchBot Hub with a cloud account"
  },
  {
    "name": "TP-Link Kasa Suite",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-tplinkkasa",
    "categories": [
      "Lighting Control"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "250502.00",
    "description": "Communication with TP-Link Kasa products"
  },
  {
    "name": "Tyba Turn 2",
    "url": "https://driverstore.rticontrol.com/driver/tyba-turn-2",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 2, 2024",
    "version": "1.03",
    "description": "Driver for Tyba Turn 2"
  },
  {
    "name": "Universal Devices ISY994 (Insteon)",
    "url": "https://driverstore.rticontrol.com/driver/universal-devices-isy994-insteon",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 26, 2014",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Vantage InFusion",
    "url": "https://driverstore.rticontrol.com/driver/vantage-infusion",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "May 8, 2020",
    "version": "1.32",
    "description": ""
  },
  {
    "name": "Vantage RFLC (Enspire)",
    "url": "https://driverstore.rticontrol.com/driver/vantage-rflc-enspire",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Oct. 28, 2011",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Vera Smart Home Control",
    "url": "https://driverstore.rticontrol.com/driver/vera-smart-home-control",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Nov. 2, 2016",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Vitrea VMax",
    "url": "https://driverstore.rticontrol.com/driver/vitrea-vmax",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 20, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Vitrea VTouch Plus Keypads",
    "url": "https://driverstore.rticontrol.com/driver/vitrea-vtouch-plus-keypads",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Dec. 30, 2015",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Vitrea VTouch Pro (formerly VBox Pro)",
    "url": "https://driverstore.rticontrol.com/driver/vitrea-vtouch-pro-formerly-vbox-pro",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 20, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "WLED Controller",
    "url": "https://driverstore.rticontrol.com/driver/wled-controller",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Sept. 15, 2023",
    "version": "1.02",
    "description": "Driver for WLED Controller"
  },
  {
    "name": "Zimi",
    "url": "https://driverstore.rticontrol.com/driver/zimi",
    "categories": [
      "Lighting Control"
    ],
    "developer": "RTI",
    "updated": "Aug. 5, 2022",
    "version": "1.7",
    "description": ""
  },
  {
    "name": "Anthem MDX-8 and MDX-16",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-anthem-mdx-8-and-mdx-16",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.0",
    "description": "The Anthem MDX RTI driver allows for control of MDX-8 and MDX-16 products."
  },
  {
    "name": "Atlona AT-CLSO",
    "url": "https://driverstore.rticontrol.com/driver/atlona-at-clso",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Feb. 25, 2016",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Atlona AT-PRO HDMI Matrix",
    "url": "https://driverstore.rticontrol.com/driver/atlona-at-pro-hdmi-matrix",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Jan. 22, 2016",
    "version": "1.30",
    "description": ""
  },
  {
    "name": "Atlona AT-UHD-CLSO-824",
    "url": "https://driverstore.rticontrol.com/driver/atlona-at-uhd-clso-824",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 11, 2016",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Atlona HDMI Switcher",
    "url": "https://driverstore.rticontrol.com/driver/atlona-hdmi-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Jan. 12, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Atlona OMNI-238 IP to Analog Audio Bridge",
    "url": "https://driverstore.rticontrol.com/driver/atlona-omni-238-ip-to-analog-audio-bridge",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 20, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Atlona OmniStream",
    "url": "https://driverstore.rticontrol.com/driver/atlona-omnistream",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 20, 2019",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Atlona Opus",
    "url": "https://driverstore.rticontrol.com/driver/atlona-opus",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 21, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Atlona Presentation Switchers",
    "url": "https://driverstore.rticontrol.com/driver/atlona-presentation-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 22, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Audac AMP523/RM523",
    "url": "https://driverstore.rticontrol.com/driver/audac-amp523-rm523",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 28, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Audac M2 Matrix Mixer",
    "url": "https://driverstore.rticontrol.com/driver/audac-m2-matrix-mixer",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 14, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Audac MTX Matrix Switches",
    "url": "https://driverstore.rticontrol.com/driver/audac-mtx-matrix-switches",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 21, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "AutoPatch PrecisLT with DSP",
    "url": "https://driverstore.rticontrol.com/driver/autopatch-precislt-with-dsp",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 22, 2021",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "Avocation Systems Mountain/Peak/Foothill Series",
    "url": "https://driverstore.rticontrol.com/driver/avocation-systems-mountain-peak-foothill-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Nov. 27, 2013",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "AVProEdge AC-CX42-AUHD",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac-cx42-auhd",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 3, 2025",
    "version": "2.8",
    "description": "AC-CX42-AUHD"
  },
  {
    "name": "AVProEdge AC-CX62-AUHD",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac-cx62-auhd",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 3, 2025",
    "version": "2.8",
    "description": "AC-CX62-AUHD"
  },
  {
    "name": "AVProEdge AC-FRESCO-44",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac-fresco-44",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Feb. 20, 2026",
    "version": "1.01",
    "description": "Driver for AVProEdge_AC-FRESCO-44"
  },
  {
    "name": "AVProEdge AC-MV-41",
    "url": "https://driverstore.rticontrol.com/driver/avpro_edge_ac-mv-41",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 24, 2026",
    "version": "3.00",
    "description": "Driver for AVPro_Edge_AC-MV-41"
  },
  {
    "name": "AVProEdge AC-MXxx Series Models",
    "url": "https://driverstore.rticontrol.com/driver/avpro-edge-ac-mxxx-series-models",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 2, 2026",
    "version": "1.17",
    "description": "Driver for AC-MXxx Series Models"
  },
  {
    "name": "AVProEdge ConferX Suite",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac-cx-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 2, 2026",
    "version": "3.02",
    "description": "Driver for AVProEdge ConferX Suite"
  },
  {
    "name": "Binary MoiP",
    "url": "https://driverstore.rticontrol.com/driver/binary-moip",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 7, 2022",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Blustream AMF41W",
    "url": "https://driverstore.rticontrol.com/driver/blustream-amf41w",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Feb. 13, 2025",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Blustream Custom Pro",
    "url": "https://driverstore.rticontrol.com/driver/blustream-custom-pro",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2020",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Blustream HDMI Matrix Switches",
    "url": "https://driverstore.rticontrol.com/driver/bluestream-av-matrix-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 1, 2021",
    "version": "2.4",
    "description": "Bluestream AV Matrix Switcher"
  },
  {
    "name": "Blustream-HMXLXXARC",
    "url": "https://driverstore.rticontrol.com/driver/blustream-hmxlxxarc",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Chowmain Software",
    "updated": "April 29, 2022",
    "version": "1.4",
    "description": ""
  },
  {
    "name": "Blustream MFP112",
    "url": "https://driverstore.rticontrol.com/driver/blustream-mfp112",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 19, 2021",
    "version": "2.3",
    "description": ""
  },
  {
    "name": "Blustream MFP62",
    "url": "https://driverstore.rticontrol.com/driver/blustream-mfp62",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "April 5, 2024",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Blustream Multicast",
    "url": "https://driverstore.rticontrol.com/driver/blustream-multicast",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Jan. 5, 2024",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Blustream Multicast IP Video (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/blustream-multicast-ip-video",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "June 21, 2020",
    "version": "1.1",
    "description": "Blustream Multicast IP Video"
  },
  {
    "name": "Blustream SW41 Series",
    "url": "https://driverstore.rticontrol.com/driver/blustream-sw41-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 15, 2023",
    "version": "1.01",
    "description": "Driver for Blustream SW41 Series"
  },
  {
    "name": "Extron MGP-641",
    "url": "https://driverstore.rticontrol.com/driver/extron-mgp-641",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Aug. 25, 2023",
    "version": "2.0",
    "description": "Driver for Extron MGP-641"
  },
  {
    "name": "Extron Scaler",
    "url": "https://driverstore.rticontrol.com/driver/extron-scaler",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Aug. 25, 2023",
    "version": "2.1",
    "description": "Driver for Extron Scaler"
  },
  {
    "name": "Gefen Toolbox",
    "url": "https://driverstore.rticontrol.com/driver/gefen-toolbox",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 4, 2013",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "HD Anywhere",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-hd-anywhere",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 24, 2025",
    "version": "1.2",
    "description": "MHUB AUDIO is a 6x4 matrix-amp with DSP EQ"
  },
  {
    "name": "HDANYWHERE MHUB PRO 2.0 (4×4)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-hdanywhere-mhub-20",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "240406.00",
    "description": "Driver for HDANYWHERE MHUB Pro 2.0 (4x4)"
  },
  {
    "name": "HDANYWHERE MHUB PRO 2.0 (8×8)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-hdanywhere-mhub-20-1",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "240406.00",
    "description": "Driver for HDANYWHERE MHUB Pro (8x8)"
  },
  {
    "name": "HDANYWHERE MHUB U (4×1+1)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-hdanywhere-mhub-20-2",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "240406.00",
    "description": "Driver for HDANYWHERE MHUB U 4x4+1"
  },
  {
    "name": "HDANYWHERE uControl Zone Processor 1 (ZP1)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-hdanywhere-mhub-20-3",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "240406.00",
    "description": "Driver for HDANYWHERE uControl Zone Processor Mini (ZP1)"
  },
  {
    "name": "HDANYWHERE uControl Zone Processor 5 (ZP5)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-hdanywhere-mhub-20-4",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 2, 2026",
    "version": "240406.00",
    "description": "Driver for HDANYWHERE uControl Zone Processor (ZP5)"
  },
  {
    "name": "Intelix FLX Matrix Switchers by Liberty AV Solutions",
    "url": "https://driverstore.rticontrol.com/driver/intelix-flx-matrix-switchers-by-liberty-av-solutions",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "May 22, 2015",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Key Digital AVoIP",
    "url": "https://driverstore.rticontrol.com/driver/key-digital-avoip",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 7, 2022",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Key Digital Matrix Switchers",
    "url": "https://driverstore.rticontrol.com/driver/key-digital-matrix-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 2, 2014",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Key Digital MS Switchers",
    "url": "https://driverstore.rticontrol.com/driver/key-digital-ms-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 1, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Knox Matrix Switchers",
    "url": "https://driverstore.rticontrol.com/driver/knox-matrix-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "May 27, 2011",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Leaf Crosspoint",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-leaf-crosspoint",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "My Device",
    "updated": "Sept. 7, 2012",
    "version": "1.0",
    "description": "Driver for Leaf Crosspoint"
  },
  {
    "name": "MuxLab Matrix Switchers",
    "url": "https://driverstore.rticontrol.com/driver/muxlab-matrix-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "June 17, 2014",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "MuxLab ProDigital",
    "url": "https://driverstore.rticontrol.com/driver/muxlab-prodigital",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 2, 2018",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "NeoPro Matrix Switchers",
    "url": "https://driverstore.rticontrol.com/driver/neopro-matrix-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Nov. 1, 2012",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Parasound Zhd Switcher",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-parasound-zhd-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "My Device",
    "updated": "July 1, 2010",
    "version": "1.0",
    "description": "Driver for Parasound Zhd Switcher"
  },
  {
    "name": "Pulse-Eight Palladia & ProAudio",
    "url": "https://driverstore.rticontrol.com/driver/pulse-eight-palladia-proaudio",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 3, 2021",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "PureLink PM Series",
    "url": "https://driverstore.rticontrol.com/driver/purelink-pm-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 19, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RGB Spectrum Linx Series",
    "url": "https://driverstore.rticontrol.com/driver/rgb-spectrum-linx-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "April 20, 2012",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "RGB Spectrum MediaWall 2900/1900",
    "url": "https://driverstore.rticontrol.com/driver/rgb-spectrum-mediawall-2900-1900",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "April 20, 2012",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "RTI VFX Flexible Matrix Switcher",
    "url": "https://driverstore.rticontrol.com/driver/rti-vfx-flexible-matrix-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Aug. 9, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "RTI VHD Matrix Switcher",
    "url": "https://driverstore.rticontrol.com/driver/rti-vhd-matrix-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VHDx",
    "url": "https://driverstore.rticontrol.com/driver/rti-vhdx",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 27, 2019",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "RTI VHX HDMI 2.0 Matrix Switcher",
    "url": "https://driverstore.rticontrol.com/driver/rti-vhx-hdmi-20-matrix-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VIP-UHD-CTRL",
    "url": "https://driverstore.rticontrol.com/driver/rti-vip-uhd-ctrl",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Dec. 6, 2024",
    "version": "1.04",
    "description": "Driver for RTI VIP-UHD-CTRL"
  },
  {
    "name": "RTI VMS-741 Multiviewer",
    "url": "https://driverstore.rticontrol.com/driver/rti-vms-741-multiviewer",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "April 19, 2019",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "RTI VMX Modular Video Distribution System",
    "url": "https://driverstore.rticontrol.com/driver/rti-vmx-modular-video-distribution-system",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 20, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VSS-51 (VSW-51 & VSW-51R) Scaling Switcher",
    "url": "https://driverstore.rticontrol.com/driver/rti-vss-51-vsw-51-vsw-51r-scaling-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "May 9, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "RTI VSW-41 HDMI 2.0 Switcher",
    "url": "https://driverstore.rticontrol.com/driver/rti-vsw-41-hdmi-20-switcher",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VWS-21T Wall Plate and Scaler",
    "url": "https://driverstore.rticontrol.com/driver/rti-vws-21t-wall-plate-and-scaler",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "March 2, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VX44 / VX88",
    "url": "https://driverstore.rticontrol.com/driver/rti-vx44-vx88",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Feb. 16, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI VXP-82",
    "url": "https://driverstore.rticontrol.com/driver/rti-vxp-82",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "June 9, 2021",
    "version": "1.01",
    "description": "VXP-82"
  },
  {
    "name": "SnapAV B-100/300/500 Series HDMatrix Switchers",
    "url": "https://driverstore.rticontrol.com/driver/snapav-b-100-300-500-series-hdmatrix-switchers",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "April 7, 2017",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Vanco Evolution",
    "url": "https://driverstore.rticontrol.com/driver/vanco-evolution",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "June 14, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Vaux Lattis Matrix",
    "url": "https://driverstore.rticontrol.com/driver/vaux-lattis-matrix",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "May 14, 2010",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Vaux Lattis Matrix V10",
    "url": "https://driverstore.rticontrol.com/driver/vaux-lattis-matrix-v10",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Nov. 23, 2010",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Video Storm CMX Series",
    "url": "https://driverstore.rticontrol.com/driver/video-storm-cmx-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 30, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Visionary Solutions PacketAV 4k",
    "url": "https://driverstore.rticontrol.com/driver/visionary-solutions-packetav-4k",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "May 24, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "WyreStorm HDMI Matrix",
    "url": "https://driverstore.rticontrol.com/driver/wyrestorm-hdmi-matrix",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Oct. 28, 2011",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Zektor Matrix",
    "url": "https://driverstore.rticontrol.com/driver/zektor-matrix",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "June 2, 2014",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Zigen HXL-44 Plus",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-zigen-hxl-44plus",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "1.03",
    "description": "Control of Zigen HXL-44Plus via IP and RS232"
  },
  {
    "name": "Zigen HXL-88 Plus",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-zigen-hxl-88plus",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "1.03",
    "description": "Control of Zigen HXL-88 Plus via IP and RS232"
  },
  {
    "name": "Zigen HXL Series",
    "url": "https://driverstore.rticontrol.com/driver/zigen-hxl-series",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "July 10, 2019",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Zigen Matrix",
    "url": "https://driverstore.rticontrol.com/driver/zigen-matrix",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "RTI",
    "updated": "Sept. 14, 2012",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Zigen SW-41 Plus",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-zigen-sw-41plus",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "1.00",
    "description": "Control of Zigen SW-41Plus 4×1 via IP and RS232"
  },
  {
    "name": "Zigen SW-42 Plus",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-zigen-sw-42plus",
    "categories": [
      "Matrix Switch"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "1.00",
    "description": "Control of Zigen SW-42Plus 4×2 via IP"
  },
  {
    "name": "Amazon Fire TV",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-amazon-fire-tv",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 28, 2026",
    "version": "260402.00",
    "description": "Driver for Amazon Fire TV"
  },
  {
    "name": "Android IP Remote",
    "url": "https://driverstore.rticontrol.com/driver/android-ip-remote",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "iNoahGuy",
    "updated": "Nov. 14, 2026",
    "version": "3.6",
    "description": "Control any Android based device over IP with zero lag."
  },
  {
    "name": "Android TV",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-zank-remote",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.7",
    "description": "Control your Android TV or Fire TV over IP with app launching"
  },
  {
    "name": "Blackmagic Hyperdeck Series",
    "url": "https://driverstore.rticontrol.com/driver/blackmagic-hyperdeck-series-driver",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "March 24, 2023",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Bose SoundTouch",
    "url": "https://driverstore.rticontrol.com/driver/bose-soundtouch",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "July 27, 2017",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Cicero Recording Software",
    "url": "https://driverstore.rticontrol.com/driver/cicero-driver",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "April 21, 2023",
    "version": "1.0",
    "description": "Driver for Cicero Recording Software"
  },
  {
    "name": "Denon HEOS (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/denon-heos",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Sept. 16, 2020",
    "version": "1.43",
    "description": "Denon HEOS"
  },
  {
    "name": "Denon iPod Dock",
    "url": "https://driverstore.rticontrol.com/driver/denon-ipod-dock",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Nov. 23, 2010",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "DF Solutions",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-df-solutions",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "My Device",
    "updated": "May 29, 2013",
    "version": "1.3",
    "description": "Driver for DF Solutions"
  },
  {
    "name": "DLO HomeDock Pro",
    "url": "https://driverstore.rticontrol.com/driver/dlo-homedock-pro",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Aug. 21, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Escient Media Server",
    "url": "https://driverstore.rticontrol.com/driver/escient-media-server",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Nov. 23, 2010",
    "version": "1.15",
    "description": ""
  },
  {
    "name": "Fusion Movie Servers",
    "url": "https://driverstore.rticontrol.com/driver/fusion-movie-servers",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "July 22, 2013",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Imerge / XiVA Media Server",
    "url": "https://driverstore.rticontrol.com/driver/imerge-xiva-media-server",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "April 30, 2010",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Integra iPod/AirPlay Dock",
    "url": "https://driverstore.rticontrol.com/driver/integra-ipod-airplay-dock",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "March 7, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Kaleidescape Media Server",
    "url": "https://driverstore.rticontrol.com/driver/kaleidescape-media-server",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "May 15, 2023",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "KEF W1 LS50W",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-kef-ls50",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "230502.00",
    "description": "Control of KEF LS50W speakers via IP"
  },
  {
    "name": "KEF W2 Media",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-kef-w2",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "July 1, 2026",
    "version": "260602.00",
    "description": "KEF W2 media player with full content navigation"
  },
  {
    "name": "KEF XIO Soundbar Media",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-kef-xio-soundbar",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "July 1, 2026",
    "version": "260602.00",
    "description": "Control of KEF XIO series device via IP"
  },
  {
    "name": "Kodi Frodo/Gotham/Helix/Krypton",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-kodi-frodo-gotham-helix-krypton",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "My Device",
    "updated": "Aug. 9, 2018",
    "version": "4.2",
    "description": "Driver for Kodi Frodo/Gotham/Helix/Krypton"
  },
  {
    "name": "Logitech Squeezebox",
    "url": "https://driverstore.rticontrol.com/driver/logitech-squeezebox",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Aug. 28, 2017",
    "version": "1.54",
    "description": ""
  },
  {
    "name": "Marantz iPod Dock",
    "url": "https://driverstore.rticontrol.com/driver/marantz-ipod-dock",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "May 28, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Marantz Network Audio Player",
    "url": "https://driverstore.rticontrol.com/driver/marantz-network-audio-player",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "April 20, 2012",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "Meridian Ellipse",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-meridian-ellipse",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "240704.00",
    "description": "Communication with Meridian Ellipse device"
  },
  {
    "name": "Naim Zone",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-naim-uniti",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "260302.00",
    "description": "Naim device audio zone control and basic meta data feedback via IP"
  },
  {
    "name": "Nuvo Wireless Audio System",
    "url": "https://driverstore.rticontrol.com/driver/nuvo-wireless-audio-system",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Nov. 15, 2019",
    "version": "1.5",
    "description": ""
  },
  {
    "name": "Ovation Music Streamers",
    "url": "https://driverstore.rticontrol.com/driver/ovation-music-streamers",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "June 24, 2022",
    "version": "22.6",
    "description": ""
  },
  {
    "name": "Pioneer Elite Network Media Player",
    "url": "https://driverstore.rticontrol.com/driver/pioneer-elite-network-media-player",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "April 20, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Popcorn Hour C200",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-popcorn-c200",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "My Device",
    "updated": "April 16, 2011",
    "version": "1.0",
    "description": "Driver for Popcorn Hour C200"
  },
  {
    "name": "Prima Cinema Player",
    "url": "https://driverstore.rticontrol.com/driver/prima-cinema-player",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Nov. 21, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Roku Media Streamer",
    "url": "https://driverstore.rticontrol.com/driver/roku-media-streamer",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Jan. 30, 2026",
    "version": "2.3",
    "description": ""
  },
  {
    "name": "Roku Pro",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-roku-pro",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "250102.00",
    "description": "Driver for Roku"
  },
  {
    "name": "RTiDock iPod Dock",
    "url": "https://driverstore.rticontrol.com/driver/rtidock-ipod-dock",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Oct. 5, 2012",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Russound MBX",
    "url": "https://driverstore.rticontrol.com/driver/russound-mbx",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "July 31, 2026",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "Sky Pro (UK)",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-skyq-pro",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "240506.00",
    "description": "Integration of SkyQ  devices with dynamic UK channel listings"
  },
  {
    "name": "SkyQ IP",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-skyq-basic",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "240506.00",
    "description": "SkyQ boxes integration and control"
  },
  {
    "name": "Sonance iPort",
    "url": "https://driverstore.rticontrol.com/driver/sonance-iport",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Dec. 9, 2011",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Sonos",
    "url": "https://driverstore.rticontrol.com/driver/sonos",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Nov. 6, 2026",
    "version": "6.3",
    "description": "Driver for Sonos"
  },
  {
    "name": "Sonos (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/sonos-media-players",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "RTI",
    "updated": "Sept. 11, 2026",
    "version": "5.15",
    "description": "Driver for Sonos"
  },
  {
    "name": "VSSL MX",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-vssl-mx",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 8, 2026",
    "version": "260402.01",
    "description": "Driver for VSSL MX Series"
  },
  {
    "name": "WD TV Live",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-wd-tv-live",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "My Device",
    "updated": "Sept. 14, 2015",
    "version": "1.3",
    "description": "Driver for WD TV Live"
  },
  {
    "name": "WiiM",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-wiim-amp",
    "categories": [
      "Media Server and Player"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "241102.00",
    "description": "Driver for WiiM WiiM amps and WiiM players"
  },
  {
    "name": "ADA Suite 8/16/8000 Series",
    "url": "https://driverstore.rticontrol.com/driver/ada-suite-8-16-8000-series",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Aug. 22, 2017",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Audac R2 Multi-zone Digital Audio Matrix",
    "url": "https://driverstore.rticontrol.com/driver/audac-r2-multi-zone-digital-audio-matrix",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "March 23, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "AudioControl Hyperion",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol_hyperion-1",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Aug. 7, 2026",
    "version": "2.00",
    "description": "Driver for AudioControl Hyperion"
  },
  {
    "name": "AudioControl Hyperion Axis",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol_axis",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "May 1, 2026",
    "version": "1.01",
    "description": "Driver for AudioControl_Axis"
  },
  {
    "name": "AudioControl Hyperion (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/audiocontrol_hyperion",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Feb. 13, 2026",
    "version": "1.01",
    "description": "Driver for AudioControl_Hyperion"
  },
  {
    "name": "AVProEdge AC AEX RC HUB",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac_aex_rc_hub",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Feb. 20, 2026",
    "version": "2.14",
    "description": "Driver for AVProEdge_AC_AEX_RC_HUB"
  },
  {
    "name": "AVProEdge AC MAX 24",
    "url": "https://driverstore.rticontrol.com/driver/avproedge_ac_max_24",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Feb. 20, 2026",
    "version": "2.14",
    "description": "Driver for AVProEdge_AC_MAX_24"
  },
  {
    "name": "AVPro Edge MXNet_10G (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/avproedge-mxnet_10g",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 3, 2025",
    "version": "2.7",
    "description": "MXNet_10G"
  },
  {
    "name": "AVPro Edge MXNet_1G (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/avproedge-mxnet_1g",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 3, 2025",
    "version": "2.7",
    "description": "MXNet_1G"
  },
  {
    "name": "AVPro Edge MXnet Dante CBOX",
    "url": "https://driverstore.rticontrol.com/driver/avpro-edge-ac-mxnet-dante-cbox",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Sept. 18, 2026",
    "version": "1.10",
    "description": "Driver for AVPro Edge AC-MXNET-DANTE-CBOX"
  },
  {
    "name": "AVPro Edge MXnet Universal",
    "url": "https://driverstore.rticontrol.com/driver/avpro-edge-mxnet-universal",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Nov. 13, 2026",
    "version": "1.07",
    "description": "Driver for AVPro Edge MXnet Universal"
  },
  {
    "name": "AVPro Edge MXNet USP (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/avpro-edge-mxnet-usp",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Feb. 13, 2026",
    "version": "3.20",
    "description": "Driver for AVPro Edge MXNet USP"
  },
  {
    "name": "B&K CT Series",
    "url": "https://driverstore.rticontrol.com/driver/bk-ct-series",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "May 27, 2011",
    "version": "1.12",
    "description": ""
  },
  {
    "name": "BluOS / Bluesound / NAD Media Players",
    "url": "https://driverstore.rticontrol.com/driver/bluos-bluesound-nad-media-players",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "Lenbrook",
    "updated": "April 24, 2026",
    "version": "2.60",
    "description": ""
  },
  {
    "name": "Blustream MV41",
    "url": "https://driverstore.rticontrol.com/driver/blustream-mv41",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Sept. 27, 2024",
    "version": "1.02",
    "description": "Driver for Blustream MV41"
  },
  {
    "name": "Denon HEOS",
    "url": "https://driverstore.rticontrol.com/driver/denon-heos-1",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Aug. 21, 2026",
    "version": "2.0",
    "description": "Driver for Denon HEOS"
  },
  {
    "name": "Elan S86AV",
    "url": "https://driverstore.rticontrol.com/driver/elan-s86av",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "June 25, 2010",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Episode Response",
    "url": "https://driverstore.rticontrol.com/driver/episode-response",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "May 20, 2022",
    "version": "1.1",
    "description": "Episode Response Amp"
  },
  {
    "name": "Klipsch KDA/Jamo JDA",
    "url": "https://driverstore.rticontrol.com/driver/klipsch-kda-jamo-jda",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 13, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Lexicon DD-8+",
    "url": "https://driverstore.rticontrol.com/driver/lexicon-dd-8",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 3, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "McIntosh Preamplifier",
    "url": "https://driverstore.rticontrol.com/driver/mcintosh-preamplifier",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 20, 2017",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Meridian 218 Zone Controller",
    "url": "https://driverstore.rticontrol.com/driver/meridian-218-zone-controller",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Oct. 6, 2017",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Monitor Audio IA Series Amplifiers",
    "url": "https://driverstore.rticontrol.com/driver/monitor-audio-ia-series-amplifiers",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Sept. 21, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "NuVo Grand Concerto",
    "url": "https://driverstore.rticontrol.com/driver/nuvo-grand-concerto",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 29, 2013",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Nuvo Simplese",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-nuvo-simplese",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "My Device",
    "updated": "Aug. 3, 2014",
    "version": "1.0",
    "description": "Driver for Nuvo Simplese"
  },
  {
    "name": "Rotel RKB Eight Channel Power Amplifier",
    "url": "https://driverstore.rticontrol.com/driver/rotel-rkb-eight-channel-power-amplifier",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Jan. 18, 2018",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "RTI AD-64",
    "url": "https://driverstore.rticontrol.com/driver/rti-ad-64",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "June 23, 2023",
    "version": "1.74",
    "description": ""
  },
  {
    "name": "RTI Miravue VIP-1",
    "url": "https://driverstore.rticontrol.com/driver/rti-miravue-vip-1",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Nov. 8, 2019",
    "version": "1.70",
    "description": ""
  },
  {
    "name": "RTI Music",
    "url": "https://driverstore.rticontrol.com/driver/rti-music",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Aug. 25, 2023",
    "version": "2.0",
    "description": "RTI MS-1 and MS-3"
  },
  {
    "name": "Russound E & C Series",
    "url": "https://driverstore.rticontrol.com/driver/russound-e-c-series",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Dec. 16, 2013",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Russound RNet Series",
    "url": "https://driverstore.rticontrol.com/driver/russound-rnet-series",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 10, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sonance C4630 / DAB-1",
    "url": "https://driverstore.rticontrol.com/driver/sonance-c4630-dab-1",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "July 10, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sonance DSP 8-130",
    "url": "https://driverstore.rticontrol.com/driver/sonance-dsp-8-130",
    "categories": [
      "Multi-Room A/V"
    ],
    "developer": "RTI",
    "updated": "Dec. 13, 2017",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "DALCNET DGM02-1248",
    "url": "https://driverstore.rticontrol.com/driver/dalcnet-dgm02-1248",
    "categories": [
      "Networking"
    ],
    "developer": "D2A",
    "updated": "Dec. 12, 2025",
    "version": "1.1",
    "description": "Driver for DALCNET DGM02-1248"
  },
  {
    "name": "Netgear M4250 Pro AV Switches",
    "url": "https://driverstore.rticontrol.com/driver/netgear-m4250-pro-av-switches",
    "categories": [
      "Networking"
    ],
    "developer": "RTI",
    "updated": "Dec. 6, 2024",
    "version": "1.02",
    "description": "Driver for Netgear M4250 Pro AV Switches"
  },
  {
    "name": "Netgear M4300 Pro AV Switches",
    "url": "https://driverstore.rticontrol.com/driver/netgear-m4300-pro-av-switches",
    "categories": [
      "Networking"
    ],
    "developer": "RTI",
    "updated": "Dec. 6, 2024",
    "version": "1.02",
    "description": "Driver for Netgear M4300 Pro AV Switches"
  },
  {
    "name": "Unifi Network",
    "url": "https://driverstore.rticontrol.com/driver/unifi-network",
    "categories": [
      "Networking"
    ],
    "developer": "RTI",
    "updated": "Dec. 5, 2025",
    "version": "1.0",
    "description": "Driver for UniFi Network products"
  },
  {
    "name": "Big Dog Power",
    "url": "https://driverstore.rticontrol.com/driver/big-dog-power",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Oct. 27, 2023",
    "version": "1.2",
    "description": "Driver for Big Dog Power"
  },
  {
    "name": "eGauge",
    "url": "https://driverstore.rticontrol.com/driver/egauge",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Aug. 30, 2013",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "GUDE EPC Series",
    "url": "https://driverstore.rticontrol.com/driver/gude-epc-series",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "April 25, 2025",
    "version": "1.0",
    "description": "Driver for GUDE EPC Series"
  },
  {
    "name": "Juice Goose IP Series Power Conditioners",
    "url": "https://driverstore.rticontrol.com/driver/juice-goose-ip-series-power-conditioners",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Nov. 22, 2019",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Luxul Power Distribution",
    "url": "https://driverstore.rticontrol.com/driver/luxul-power-distribution",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "May 24, 2019",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "LynTec RPC",
    "url": "https://driverstore.rticontrol.com/driver/lyntec-rpc",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "May 22, 2026",
    "version": "1.0",
    "description": "Driver for LynTec RPC"
  },
  {
    "name": "Middle Atlantic Racklink",
    "url": "https://driverstore.rticontrol.com/driver/middle-atlantic-racklink",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "April 20, 2018",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Netio Pro",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-netio-rest-pro-edition",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 31, 2026",
    "version": "260202.00",
    "description": "Netio REST (Pro Edition)"
  },
  {
    "name": "Panamax Other Models",
    "url": "https://driverstore.rticontrol.com/driver/panamax-other-models",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Oct. 29, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Panamax VT1512",
    "url": "https://driverstore.rticontrol.com/driver/panamax-vt1512",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Oct. 29, 2021",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Power Tracker",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-power-tracker",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "My Device",
    "updated": "Sept. 7, 2012",
    "version": "1.0",
    "description": "Driver for Power Tracker"
  },
  {
    "name": "SurgeX Axess",
    "url": "https://driverstore.rticontrol.com/driver/surgex-axess",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Aug. 28, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "SurgeX Squid",
    "url": "https://driverstore.rticontrol.com/driver/surgex-squid",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "March 3, 2023",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "TED 5000",
    "url": "https://driverstore.rticontrol.com/driver/ted-5000",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "March 16, 2012",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "TPLink",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-tplink",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "1.2",
    "description": "The TP-Link RTI driver allows for control over a wide range of TP-Link WiFi products."
  },
  {
    "name": "TP-Link Kasa Plug",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-tplinkkasa-plug",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "250102.00",
    "description": "Communication with TP-Link Kasa products, inc. WI-FI Smart Plugs and Power Strips"
  },
  {
    "name": "WattBox",
    "url": "https://driverstore.rticontrol.com/driver/wattbox",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Sept. 14, 2018",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "WattBox 150/250/800",
    "url": "https://driverstore.rticontrol.com/driver/wattbox-150-250-800",
    "categories": [
      "Power and Rack Management"
    ],
    "developer": "RTI",
    "updated": "Oct. 15, 2021",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "AJAX Jeweller Series",
    "url": "https://driverstore.rticontrol.com/driver/ajaxsys",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "March 3, 2023",
    "version": "1.0",
    "description": "Ajax Jeweller Security Systems"
  },
  {
    "name": "DSC Power Series",
    "url": "https://driverstore.rticontrol.com/driver/dsc-power-series",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "July 15, 2022",
    "version": "1.06",
    "description": ""
  },
  {
    "name": "Elk E27",
    "url": "https://driverstore.rticontrol.com/driver/elk-e27",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Oct. 23, 2026",
    "version": "1.0",
    "description": "Driver for Elk E27"
  },
  {
    "name": "ELK/Ness M1",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-elk-ness-m1",
    "categories": [
      "Security"
    ],
    "developer": "My Device",
    "updated": "July 2, 2011",
    "version": "1.03",
    "description": "Driver for ELK/Ness M1"
  },
  {
    "name": "ELK / Ness Products",
    "url": "https://driverstore.rticontrol.com/driver/elk-ness-products",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Jan. 6, 2023",
    "version": "1.07",
    "description": ""
  },
  {
    "name": "GE Concord",
    "url": "https://driverstore.rticontrol.com/driver/ge-concord",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "July 20, 2017",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "HAI OmniPro II",
    "url": "https://driverstore.rticontrol.com/driver/hai-omnipro-ii",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Aug. 19, 2016",
    "version": "1.22",
    "description": ""
  },
  {
    "name": "Hikvision DS-K2604",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-hikvision-ds-k2604",
    "categories": [
      "Security"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.1",
    "description": "Control the Hikvision DS-K2604 Door controler over IP. Full control and status of all four doors."
  },
  {
    "name": "Honeywell 4232CBM",
    "url": "https://driverstore.rticontrol.com/driver/honeywell-4232cbm",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "June 6, 2019",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Honeywell EnvisaLink",
    "url": "https://driverstore.rticontrol.com/driver/honeywell-envisalink",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Nov. 9, 2018",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Honeywell Vista Security Panel",
    "url": "https://driverstore.rticontrol.com/driver/honeywell-vista-security-panel",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "July 20, 2017",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Inner Range Inception",
    "url": "https://driverstore.rticontrol.com/driver/inner-range-inception",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Jan. 2, 2025",
    "version": "1.8",
    "description": "Driver for Inner Range Inception"
  },
  {
    "name": "Lilin NVR",
    "url": "https://driverstore.rticontrol.com/driver/lilin-nvr",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Dec. 8, 2023",
    "version": "1.0",
    "description": "Driver for Lilin NVR"
  },
  {
    "name": "Lilin PTZ Cameras",
    "url": "https://driverstore.rticontrol.com/driver/lilin-ptz-cameras",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Dec. 8, 2023",
    "version": "1.0",
    "description": "Driver for Lilin PTZ Cameras"
  },
  {
    "name": "Unifi Protect",
    "url": "https://driverstore.rticontrol.com/driver/unifi-protect",
    "categories": [
      "Security"
    ],
    "developer": "RTI",
    "updated": "Dec. 5, 2025",
    "version": "1.0",
    "description": "Driver for Unifi Protect products"
  },
  {
    "name": "Dahua DVR",
    "url": "https://driverstore.rticontrol.com/driver/dahua-dvr",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "Sept. 10, 2021",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "ICRealtime DVR",
    "url": "https://driverstore.rticontrol.com/driver/icrealtime-dvr",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "Sept. 3, 2021",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "Illumivue Surveillance",
    "url": "https://driverstore.rticontrol.com/driver/illumivue-surveillance",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "Oct. 2, 2026",
    "version": "1.0",
    "description": "Driver for Illumivue Surveillance"
  },
  {
    "name": "Lux NVR",
    "url": "https://driverstore.rticontrol.com/driver/lux-nvr",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "Jan. 19, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Visca PTZ Cameras",
    "url": "https://driverstore.rticontrol.com/driver/visca",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "June 21, 2024",
    "version": "1.02",
    "description": "Driver for Visca PTZ Cameras"
  },
  {
    "name": "Wirepath DVR",
    "url": "https://driverstore.rticontrol.com/driver/wirepath-dvr",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "May 14, 2014",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Wirepath PTZ IP Camera",
    "url": "https://driverstore.rticontrol.com/driver/wirepath-ptz-ip-camera",
    "categories": [
      "Surveillance"
    ],
    "developer": "RTI",
    "updated": "Aug. 15, 2014",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "Asterisk",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-asterisk",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "My Device",
    "updated": "June 17, 2015",
    "version": "1.1",
    "description": "Driver for Asterisk"
  },
  {
    "name": "Barco Clickshare",
    "url": "https://driverstore.rticontrol.com/driver/barco-clickshare",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "March 24, 2023",
    "version": "1.21",
    "description": ""
  },
  {
    "name": "Cisco Phone Menu",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-cisco-phone-menu",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "My Device",
    "updated": "Oct. 24, 2012",
    "version": "2.2",
    "description": "Driver for Cisco Phone Menu"
  },
  {
    "name": "Cisco Telepresence",
    "url": "https://driverstore.rticontrol.com/driver/cisco-telepresence",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "May 21, 2015",
    "version": "1.11",
    "description": ""
  },
  {
    "name": "Elmo PX Series Driver",
    "url": "https://driverstore.rticontrol.com/driver/elmo-px-series-driver",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "March 24, 2023",
    "version": "1.01",
    "description": "Driver for Elmo PX Series Driver"
  },
  {
    "name": "LifeSize",
    "url": "https://driverstore.rticontrol.com/driver/lifesize",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "July 6, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "LifeSize Icon",
    "url": "https://driverstore.rticontrol.com/driver/lifesize-icon",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Dec. 7, 2017",
    "version": "1.27",
    "description": ""
  },
  {
    "name": "Polycom HDX",
    "url": "https://driverstore.rticontrol.com/driver/polycom-hdx",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Feb. 19, 2013",
    "version": "1.02",
    "description": ""
  },
  {
    "name": "PTZOptics PT12X-SDI",
    "url": "https://driverstore.rticontrol.com/driver/ptzoptics-pt12x-sdi",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Sept. 21, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI UC-IVB-50",
    "url": "https://driverstore.rticontrol.com/driver/rti-uc-ivb-50",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Feb. 15, 2024",
    "version": "1.0",
    "description": "Driver for RTI UC-IVB-50"
  },
  {
    "name": "RTI VXU-SC Soft Codec Video Conferencing System",
    "url": "https://driverstore.rticontrol.com/driver/rti-vxu-sc-soft-codec-video-conferencing-system",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Feb. 15, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Vaddio ClearVIEW",
    "url": "https://driverstore.rticontrol.com/driver/vaddio-clearview",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Oct. 5, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Vaddio IntelliSHOT",
    "url": "https://driverstore.rticontrol.com/driver/vaddio-intellishot",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "Nov. 13, 2026",
    "version": "1.0",
    "description": "Driver for Vaddio IntelliSHOT"
  },
  {
    "name": "Vaddio Roboshot",
    "url": "https://driverstore.rticontrol.com/driver/vaddio-roboshot",
    "categories": [
      "Teleconferencing"
    ],
    "developer": "RTI",
    "updated": "March 25, 2022",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "RTI Virtual Alarm",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-alarm",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Alarm"
  },
  {
    "name": "RTI Virtual AV Receiver",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-av-receiver",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual AV Receiver"
  },
  {
    "name": "RTI Virtual Blu-ray",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-blu-ray",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Blu-ray"
  },
  {
    "name": "RTI Virtual Climate",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-climate",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Climate"
  },
  {
    "name": "RTI Virtual HDMI Matrix",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-hdmi-matrix",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual HDMI Matrix"
  },
  {
    "name": "RTI Virtual Lighting",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-lighting",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Lighting"
  },
  {
    "name": "RTI Virtual Multiroom Amp",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-multiroom-amp",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Multiroom Amp"
  },
  {
    "name": "RTI Virtual Music Streamer",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-music-streamer",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Music Streamer"
  },
  {
    "name": "RTI Virtual Satellite Radio",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-satellite-radio",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Satellite Radio"
  },
  {
    "name": "RTI Virtual Satellite Receiver",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-satellite-receiver",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Satellite Receiver"
  },
  {
    "name": "RTI Virtual Shades",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-shades",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual Shades"
  },
  {
    "name": "RTI Virtual Soundbar",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-soundbar",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "Jan. 24, 2025",
    "version": "1.0",
    "description": "Driver for RTI Virtual Soundbar"
  },
  {
    "name": "RTI Virtual TV",
    "url": "https://driverstore.rticontrol.com/driver/rti-virtual-tv",
    "categories": [
      "Training"
    ],
    "developer": "RTI",
    "updated": "March 8, 2024",
    "version": "1.0",
    "description": "Driver for RTI Virtual TV"
  },
  {
    "name": "Contemporary Research 232-ATSC+",
    "url": "https://driverstore.rticontrol.com/driver/contemporary-research-232-atsc",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "July 6, 2012",
    "version": "0.9",
    "description": ""
  },
  {
    "name": "DirecTV",
    "url": "https://driverstore.rticontrol.com/driver/directv",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "Jan. 17, 2019",
    "version": "1.32",
    "description": ""
  },
  {
    "name": "Dish Network",
    "url": "https://driverstore.rticontrol.com/driver/dish-network",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "Dec. 16, 2022",
    "version": "1.0",
    "description": "Dish Network Hopper"
  },
  {
    "name": "Marantz Tuner ST7001",
    "url": "https://driverstore.rticontrol.com/driver/marantz-tuner-st7001",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "June 3, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "NuVo NV-T2SIR Tuner",
    "url": "https://driverstore.rticontrol.com/driver/nuvo-nv-t2sir-tuner",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "Sept. 22, 2011",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Parasound Ztuner",
    "url": "https://driverstore.rticontrol.com/driver/parasound-ztuner",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "June 3, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sirius Tuner",
    "url": "https://driverstore.rticontrol.com/driver/sirius-tuner",
    "categories": [
      "Tuner"
    ],
    "developer": "RTI",
    "updated": "March 18, 2011",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Calculator Utility",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-calculator",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "Aug. 31, 2012",
    "version": "1.1",
    "description": "Driver for Calculator"
  },
  {
    "name": "Clock Driver",
    "url": "https://driverstore.rticontrol.com/driver/clock-driver",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Feb. 25, 2013",
    "version": "3.41",
    "description": ""
  },
  {
    "name": "Email / SMTP Client",
    "url": "https://driverstore.rticontrol.com/driver/email-smtp-client",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "March 27, 2017",
    "version": "2.52",
    "description": ""
  },
  {
    "name": "Generic TCP Server",
    "url": "https://driverstore.rticontrol.com/driver/clever-av-generic-tcp-server",
    "categories": [
      "Utility"
    ],
    "developer": "Clever AV Solutions",
    "updated": "May 26, 2026",
    "version": "1.0",
    "description": "Generic TCP Server Driver"
  },
  {
    "name": "Global Cache",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-global-cache",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "April 19, 2014",
    "version": "2.1",
    "description": "Driver for Global Cache"
  },
  {
    "name": "Global Cache Model",
    "url": "https://driverstore.rticontrol.com/driver/global-cache-model",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Feb. 10, 2012",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Home Butler",
    "url": "https://driverstore.rticontrol.com/driver/home-butler",
    "categories": [
      "Utility"
    ],
    "developer": "iNoahGuy",
    "updated": "Nov. 15, 2025",
    "version": "3.13",
    "description": "Add Voice control to RTI"
  },
  {
    "name": "HTTP Commander",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-http-commander",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "June 25, 2026",
    "version": "1.50",
    "description": "Driver for HTTP to Processor Control"
  },
  {
    "name": "HTTP Server Driver",
    "url": "https://driverstore.rticontrol.com/driver/rti-http-server-driver",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "May 29, 2020",
    "version": "1.4",
    "description": "RTI HTTP Server Driver"
  },
  {
    "name": "Hunter Douglas PowerView Gen3",
    "url": "https://driverstore.rticontrol.com/driver/powerview-gen3",
    "categories": [
      "Utility"
    ],
    "developer": "Control Concepts, Inc.",
    "updated": "Oct. 14, 2022",
    "version": "1.2",
    "description": "PowerView Gen3"
  },
  {
    "name": "HyperTerminal",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-hyperterminal",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "April 1, 2026",
    "version": "1.01",
    "description": "Driver for RTI HyperTerminal"
  },
  {
    "name": "ID11 Addon Driver",
    "url": "https://driverstore.rticontrol.com/driver/clever-av-id11-addon",
    "categories": [
      "Utility"
    ],
    "developer": "Clever AV Solutions",
    "updated": "Dec. 23, 2025",
    "version": "1.03",
    "description": ""
  },
  {
    "name": "Integer Helper",
    "url": "https://driverstore.rticontrol.com/driver/integer-helper",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Aug. 7, 2020",
    "version": "1.3",
    "description": ""
  },
  {
    "name": "Intrinsic Multicast Driver",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-intrinsic-multicast-driver",
    "categories": [
      "Utility"
    ],
    "developer": "Intrinsic Dev",
    "updated": "March 30, 2026",
    "version": "250602.00",
    "description": "Provides discovery and communication between drivers."
  },
  {
    "name": "iPort Surface Mount with Buttons",
    "url": "https://driverstore.rticontrol.com/driver/iport-surface-mount-with-buttons",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "July 29, 2016",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Kramer 2000",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-kramer-2000",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "Oct. 10, 2012",
    "version": "1.1",
    "description": "Driver for Kramer 2000"
  },
  {
    "name": "Local Flags",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-local-flags",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "Feb. 28, 2026",
    "version": "1.11",
    "description": "Driver for RTI LocalFlags"
  },
  {
    "name": "Logic Helper",
    "url": "https://driverstore.rticontrol.com/driver/logic-helper",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "July 2, 2021",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Memory Watcher",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-memory-watcher",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "June 16, 2015",
    "version": "1.0",
    "description": "Driver for Memory Watcher"
  },
  {
    "name": "One Touch Button",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-one-touch-button",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "Jan. 11, 2014",
    "version": "1.1",
    "description": "Driver for One Touch Button"
  },
  {
    "name": "Prayer Time",
    "url": "https://driverstore.rticontrol.com/driver/prayer-time",
    "categories": [
      "Utility"
    ],
    "developer": "D2A",
    "updated": "Feb. 6, 2026",
    "version": "1.0",
    "description": "Driver for Prayer Time"
  },
  {
    "name": "Print To Log",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-print-to-log",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "Dec. 8, 2025",
    "version": "1.00",
    "description": "Driver for RTI Print To Log"
  },
  {
    "name": "Processor to Processor Communication",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-proc-to-proc-communication",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "April 1, 2026",
    "version": "2.00",
    "description": "Driver for Comms Between RTI Processors"
  },
  {
    "name": "Pronestor Display",
    "url": "https://driverstore.rticontrol.com/driver/pronestor-display",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Dec. 12, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Pushover",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-pushover",
    "categories": [
      "Utility"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "3.2",
    "description": "The Pushover driver can be used to send and receive notifications from the Pushover service."
  },
  {
    "name": "QNAP Management",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-qnap-management",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "April 13, 2012",
    "version": "1.0",
    "description": "Driver for QNAP Management"
  },
  {
    "name": "ReCalc",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-recalc-v220",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "April 1, 2026",
    "version": "2.20",
    "description": "Driver for Coupe Automation XP Processors"
  },
  {
    "name": "RTI System Macro Triggers",
    "url": "https://driverstore.rticontrol.com/driver/rti-system-macro-triggers",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "June 3, 2009",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "RTI XP Diagnostics",
    "url": "https://driverstore.rticontrol.com/driver/xp-diagnostics",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Dec. 15, 2023",
    "version": "1.2",
    "description": "XP Diagnostics"
  },
  {
    "name": "Schedule Manager",
    "url": "https://driverstore.rticontrol.com/driver/schedule-manager",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "July 17, 2026",
    "version": "2.1",
    "description": "Driver for Schedule Manager"
  },
  {
    "name": "Shabbat Utility",
    "url": "https://driverstore.rticontrol.com/driver/shabbat-utility",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "July 24, 2026",
    "version": "1.13",
    "description": "Driver for Shabbat Utility"
  },
  {
    "name": "Shelly Sense",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-shellyplussense",
    "categories": [
      "Utility"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "240404.00",
    "description": "Monitor and feedback of a Shelly Plus i4 device"
  },
  {
    "name": "Simple Logic",
    "url": "https://driverstore.rticontrol.com/driver/coupe-automation-simple-logic",
    "categories": [
      "Utility"
    ],
    "developer": "Coupe Automation",
    "updated": "April 1, 2026",
    "version": "1.21",
    "description": "Driver for Coupe Automation XP Processors"
  },
  {
    "name": "System To System",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-system-to-system",
    "categories": [
      "Utility"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.0",
    "description": "Join different RTI processors together to send boolean, integer and string values and also trigger events."
  },
  {
    "name": "Text Helper",
    "url": "https://driverstore.rticontrol.com/driver/text-helper",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Aug. 7, 2020",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Two-way Strings",
    "url": "https://driverstore.rticontrol.com/driver/rti-two-way-strings",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "July 31, 2020",
    "version": "2.7",
    "description": "RTI Two-way Strings"
  },
  {
    "name": "Two Way Strings",
    "url": "https://driverstore.rticontrol.com/driver/intrinsic-intrinsicdev-tcp-two-way-strings-driver",
    "categories": [
      "Utility"
    ],
    "developer": "Intrinsic Dev",
    "updated": "April 1, 2026",
    "version": "240302.00",
    "description": "Send up to 50 custom strings as normal or hexadecimal text via TCP"
  },
  {
    "name": "Virtual Presets",
    "url": "https://driverstore.rticontrol.com/driver/virtual-presets",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Feb. 5, 2013",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Voice Viper",
    "url": "https://driverstore.rticontrol.com/driver/voice-viper",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "April 9, 2014",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "WeMo",
    "url": "https://driverstore.rticontrol.com/driver/mydevice-wemo",
    "categories": [
      "Utility"
    ],
    "developer": "My Device",
    "updated": "Oct. 26, 2013",
    "version": "1.1",
    "description": "Driver for WeMo"
  },
  {
    "name": "WOL Magic Packet",
    "url": "https://driverstore.rticontrol.com/driver/wol-magic-packet",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Nov. 16, 2018",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "YJ-TMSv3 Tank Monitor System",
    "url": "https://driverstore.rticontrol.com/driver/yj-tmsv3-tank-monitor-system",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Feb. 23, 2024",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "Zigbee Network Test",
    "url": "https://driverstore.rticontrol.com/driver/rti-zigbee-network-test",
    "categories": [
      "Utility"
    ],
    "developer": "RTI",
    "updated": "Sept. 27, 2019",
    "version": "1.12",
    "description": "RTI Zigbee Network Test"
  },
  {
    "name": "Barco Pulse Platform (Legacy)",
    "url": "https://driverstore.rticontrol.com/driver/barco-pulse-platform",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "July 2, 2021",
    "version": "1.32",
    "description": "Barco Pulse Platform"
  },
  {
    "name": "Digital Projection",
    "url": "https://driverstore.rticontrol.com/driver/digital-projection",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Nov. 15, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Epson PowerLite",
    "url": "https://driverstore.rticontrol.com/driver/epson-powerlite",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "May 22, 2026",
    "version": "1.1",
    "description": ""
  },
  {
    "name": "Epson Q-Series",
    "url": "https://driverstore.rticontrol.com/driver/epson-q-series",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Nov. 13, 2026",
    "version": "1.0",
    "description": "Driver for Epson Q-Series"
  },
  {
    "name": "Epson Universal Projector",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-epson-universal-projector",
    "categories": [
      "Video Projector"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.2",
    "description": "Universal Epson ESC/VP projector control with advaced control and automation features"
  },
  {
    "name": "Hitachi CP Series Projector",
    "url": "https://driverstore.rticontrol.com/driver/hitachi-cp-series-projector",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Feb. 15, 2019",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "JVC D-ILA Projector",
    "url": "https://driverstore.rticontrol.com/driver/jvc-d-ila-projector",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "May 1, 2026",
    "version": "1.51",
    "description": ""
  },
  {
    "name": "JVC LX Series DLP",
    "url": "https://driverstore.rticontrol.com/driver/jvc-lx-series-dlp",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "July 2, 2021",
    "version": "1.00",
    "description": ""
  },
  {
    "name": "Optoma Universal Projector",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-optoma-universal-projector",
    "categories": [
      "Video Projector"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "2.0",
    "description": "Universal Optoma Projector control for RTI"
  },
  {
    "name": "PJLink",
    "url": "https://driverstore.rticontrol.com/driver/pjlink",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Oct. 11, 2018",
    "version": "1.01",
    "description": ""
  },
  {
    "name": "PJLink",
    "url": "https://driverstore.rticontrol.com/driver/chowmain-pjlink",
    "categories": [
      "Video Projector"
    ],
    "developer": "Chowmain Software",
    "updated": "Oct. 23, 2025",
    "version": "3.1",
    "description": "The PJLink driver gives you control over PJLink enabled projectors"
  },
  {
    "name": "Screen Research MCS",
    "url": "https://driverstore.rticontrol.com/driver/screen-research-mcs",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Oct. 18, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sim2 Nero 4",
    "url": "https://driverstore.rticontrol.com/driver/sim2-nero-4",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Sept. 27, 2019",
    "version": "1.0",
    "description": ""
  },
  {
    "name": "Sony 4k Projector",
    "url": "https://driverstore.rticontrol.com/driver/sony-4k-projector-2019-and-newer",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "Sept. 12, 2022",
    "version": "1.04",
    "description": "Sony 4k Projector - 2019 and Newer"
  },
  {
    "name": "Sony VPL Video Projectors",
    "url": "https://driverstore.rticontrol.com/driver/sony-vpl-video-projectors",
    "categories": [
      "Video Projector"
    ],
    "developer": "RTI",
    "updated": "April 11, 2014",
    "version": "1.2",
    "description": ""
  },
  {
    "name": "V-HOME Smart Projector",
    "url": "https://driverstore.rticontrol.com/driver/v-home-smart-projector",
    "categories": [
      "Video Projector"
    ],
    "developer": "Hisense USA Corporation",
    "updated": "Nov. 17, 2026",
    "version": "1.0",
    "description": "Driver for V-HOME Smart Projector"
  }
]

# pdf = FPDF()
# pdf.add_page()

# pdf.set_font("Helvetica", size=9)

# doc = Document()
# doc.add_heading("Drivers List", level=1)

# count=0;
# for i in a:
#     count=count+1
#     line = f"{str(count)+ " . " + i["name"] + " - " + i["url"]}"
#     doc.add_paragraph(line)
    # pdf.cell(0, 10, text=line, new_x="LMARGIN", new_y="NEXT")

# pdf.output("drivers.pdf")
# doc.save("drivers.docx")
# print("created successfully")
dri_names= ['2N Doorbell', 'DoorBird', 'DoorBird (Legacy)', 'Genie & Overhead Door', 'ismartgate Garage', 'MyQ', 'RTI IVDS', 'Tailwind', 'AudioControl BijouSeries', 'AudioControl CM-Series (Legacy)', 'AudioControl CMx-750 and Director Series', 'AudioControl DirectorSeries (Legacy)', 'AudioControl RS Series', 'Audio Technica ATDM-0604a', 'Australian Monitor ZoneMix', 'Blustream NPA100DA', 'Calypso CA-1000', 'Ecler CA Series Amplifier', 'Klipsch C-Series Subwoofers', 'Rotel RC Series Preamplifier', 'RTI AD DSP Matrix', 'RTI AMR-220 Amplifier', 'RTI AMR-350 3 Input Amplifier', 'RTI AMV-340 3 Input Constant Voltage Amplifier with Mic', 'RTI CP Amplifier', 'RTI Multi-Z AMP', 'Stewart Audio AV-25', 'Trinnov Altitude', 'Velodyne Digital Drive', 'Velodyne SubContractor', 'WiiM', 'ADA Suite 7.1', 'Anthem AVM 50v Processor', 'Anthem MRX Receivers', 'Anthem MRX x10 Series', 'Anthem MRXx40', 'Arcam AVR', 'Arcam AVR390/550/850/860', 'Arcam AVR Series', 'AudioControl AVR', 'AudioControl AVR-7/9', 'Bose Lifetyle Home Entertainment System', 'Cary Audio Cinema 11a', 'Classe Audio Surround Processor', 'Classe Sigma SSP Processor', 'Datasat AP20', 'Denon Marantz AVR', 'DVDO iScan Duo', 'Dynaudio Focus', 'Integra/Onkyo/Pioneer Receiver  2017+', 'Integra Receiver DTR (Deprecated)', 'JBL Synthesis SDP-45', 'Klipsch Flexus', 'Marantz Receiver SR', 'Meridian Surround Controller', 'NAD Receiver', 'Novastar H Series (Lite)', 'Parasound ZPre2', 'Parsound Zpre2 Zone Preamplifier', 'Pioneer Elite SC', 'Rotel RAP1580 Surround Amplified Processor', 'Rotel RC-1570 Stereo Preamplifier', 'Rotel Receiver', 'Rotel RSP1582 Sound Processor', 'Samsung Soundbars', 'Sony ES Receiver', 'Sony Sound Bars', 'Sony STR-ZA', 'Steinway-Lyngdorf Model P200 Surround Processor', 'Steinway-Lyngdorf Surround Processor', 'Storm Audio ISP', 'Yamaha AVENTAGE Receiver', 'Yamaha MusicCast', 'Yamaha MusicCast (Legacy)', 'Yamaha Receiver RX-V', 'Yamaha Receiver RX-Z11', 'Advantage Air', 'Advantage Air Gen III', 'Aprilaire', 'Aprilaire WiFi Thermostats', 'CoolMaster', 'CoolMasterNet', 'Dooya', 'Ecobee Thermostats', 'Hayward OmniLogic', 'Heatmiser Neo', 'Heatmiser NeoHub Pro', 'Jandy AquaLink RS', 'Jandy iAquaLink', 'Netatmo Weather Station', 'Network Thermostat', 'Pentair IntelliTouch', 'Polyaire Airtouch5', 'Radio Thermostat Company', 'RTI Works with Nest', 'Sensibo', 'Tekmar', 'Venstar ColorTouch', 'Weather', 'Denon Blu-Ray Player DBP-2012UD', 'Integra Blu-Ray/DVD', 'Magnetar Blu-ray Players', 'OPPO Bluray', 'Pioneer BDP-LX52', 'Pioneer Elite BDP', 'Pioneer UDP-LX Blu-ray Player', 'Samsung IP UHD Bluray', 'Sharp BD-AMS20U Bluray', 'Sony BDP-CX7000ES', 'Aquavision Nexus Televisions', 'Aquavision Television', 'Dynamic Camera Feed', 'Everpass Control API', 'Hisense Commercial Display', 'Hisense Smart TV', 'Hisense V-HOME Smart TV', 'Just Video Walls H Series', 'Just Video Walls M Series', 'LG IP', 'LG IP', 'LG Television', 'Loewe sl8 Series', 'MAXHUB CMA and UW Series', 'NEC_P402 Display', 'Panasonic IP TV', 'Panasonic Pro Plasma Displays', 'Panasonic Viera', 'Philips BDL TV', 'Samsung Commercial Displays', 'Samsung Ex-Link', 'Samsung IP', 'Samsung IP Television', 'Samsung Tizen OS TV with SmartThings', 'Sharp Elite', 'SkyQ Lite', 'SMART Board 6000S Series', 'Sony Bravia IP', 'Sony Bravia IP', 'Sony Bravia IP (Non-PSK Models)', 'Sony Bravia Simple IP', 'Sony Bravia Simple IP', 'Sony BRAVIA with Google TV', 'Sony XBR Series LCD', 'Sunbrite Solis IP', 'SunBrite TV', 'SunBrite Veranda 3', 'Tizen OS for Samsung TVs', 'Allen and Heath AHM', 'Atlas Atmosphere', 'AtlasIED Atmosphere (Deprecated)', 'Beale Street Audio BAV', 'Biamp Nexia', 'Biamp Tesira', 'Bose ControlSpace', 'BSS Audio London BLU v1', 'BSS Audio London BLU v2', 'ClearOne Converge', 'ClearOne CONVERGE Pro 2', 'Cloud DCM-1/e', 'Digital View DRR-2T DAB', 'Ecler DAM614', 'Ecler eMIMO1616', 'Ecler MIM01212SG', 'Ecler MIMO4040DN', 'Ecler MIMO88SG', 'Ecler VEO-XRI1C', 'Ecler VEO-XRI2L', 'LEA Connect 168D', 'Polycom SoundStructure', 'PTN PA2 / ISLink DPA-22', 'QSC MP-M', 'QSC Q-SYS', 'Shure P300', 'Symetrix', 'TOA M-9000M2 Pre-Amplifier', 'Yamaha MRX', 'Yamaha MTX', 'Yamaha YSP-4300', 'Digital View Set Top', 'Humax', 'Reavon UBR Series', 'TiVo', 'TiVo Series 3 / HD', 'Xbox One Power', 'Aquameta', 'Irrigation Caddy', 'Rachio Iro Irrigation Controller', 'Rain8net Irrigation Controller', '3 Day Blinds', 'Black Nova Keypads', 'CentraLite JetStream', 'Clipsal C-Bus', 'Clipsal C-Bus', 'Converging Systems Intelligent Lighting Controller', 'Cooper Wiring Devices Aspire RF', 'Current Products Corp Curtains', 'Dynalite', 'Ezlo Smart Home Hubs', 'Fibaro HC3', 'FXLuminaire Luxor', 'Interactive Technologies CueServer 2', 'Legrand Shading Systems QzHub', 'Legrand Shading Systems QzHub3', 'Leviton ViziaRF', 'LightwaveRF Lighting', 'LiteTouch 5000LC', 'Lutron Caseta / RA2 Select', 'Lutron Grafik Eye', 'Lutron Homeworks', 'Lutron Homeworks RF', 'Lutron LEAP', 'Lutron LEAP (Legacy)', 'Lutron myRoom', 'Lutron RadioRA', 'Lutron RadioRA 2 / Homeworks QS', 'Lutron Vive', 'MotionBlinds', 'Nice DMBM', 'NuLED SpectraDrive', 'On-Q RFLC Lighting', 'Orro Lighting', 'Philips Hue', 'Philips Hue Lighting (Deprecated)', 'Philips Hue Lighting System (v2.X)', 'PowerShades', 'QMotion QConnect', 'Qmotion QIS', 'Qmotion Qsync', 'Qmotion Qube', 'QMotion Shades', 'Rako Lighting', 'Rako Wireless', 'Rollease Pulse 2', 'Rollease Pulse Hub', 'RTI KNX IP Driver', 'Shelly', 'Silent Gliss Shades', 'Somfy Connect UAI Plus', 'Somfy ILT', 'Somfy myLink', 'Somfy POE Motor', 'Somfy TaHoma (EU)', 'Somfy TaHoma North America', 'SwitchBot', 'TP-Link Kasa Suite', 'Tyba Turn 2', 'Universal Devices ISY994 (Insteon)', 'Vantage InFusion', 'Vantage RFLC (Enspire)', 'Vera Smart Home Control', 'Vitrea VMax', 'Vitrea VTouch Plus Keypads', 'Vitrea VTouch Pro (formerly VBox Pro)', 'WLED Controller', 'Zimi', 'Anthem MDX-8 and MDX-16', 'Atlona AT-CLSO', 'Atlona AT-PRO HDMI Matrix', 'Atlona AT-UHD-CLSO-824', 'Atlona HDMI Switcher', 'Atlona OMNI-238 IP to Analog Audio Bridge', 'Atlona OmniStream', 'Atlona Opus', 'Atlona Presentation Switchers', 'Audac AMP523/RM523', 'Audac M2 Matrix Mixer', 'Audac MTX Matrix Switches', 'AutoPatch PrecisLT with DSP', 'Avocation Systems Mountain/Peak/Foothill Series', 'AVProEdge AC-CX42-AUHD', 'AVProEdge AC-CX62-AUHD', 'AVProEdge AC-FRESCO-44', 'AVProEdge AC-MV-41', 'AVProEdge AC-MXxx Series Models', 'AVProEdge ConferX Suite', 'Binary MoiP', 'Blustream AMF41W', 'Blustream Custom Pro', 'Blustream HDMI Matrix Switches', 'Blustream-HMXLXXARC', 'Blustream MFP112', 'Blustream MFP62', 'Blustream Multicast', 'Blustream Multicast IP Video (Legacy)', 'Blustream SW41 Series', 'Extron MGP-641', 'Extron Scaler', 'Gefen Toolbox', 'HD Anywhere', 'HDANYWHERE MHUB PRO 2.0 (4×4)', 'HDANYWHERE MHUB PRO 2.0 (8×8)', 'HDANYWHERE MHUB U (4×1+1)', 'HDANYWHERE uControl Zone Processor 1 (ZP1)', 'HDANYWHERE uControl Zone Processor 5 (ZP5)', 'Intelix FLX Matrix Switchers by Liberty AV Solutions', 'Key Digital AVoIP', 'Key Digital Matrix Switchers', 'Key Digital MS Switchers', 'Knox Matrix Switchers', 'Leaf Crosspoint', 'MuxLab Matrix Switchers', 'MuxLab ProDigital', 'NeoPro Matrix Switchers', 'Parasound Zhd Switcher', 'Pulse-Eight Palladia & ProAudio', 'PureLink PM Series', 'RGB Spectrum Linx Series', 'RGB Spectrum MediaWall 2900/1900', 'RTI VFX Flexible Matrix Switcher', 'RTI VHD Matrix Switcher', 'RTI VHDx', 'RTI VHX HDMI 2.0 Matrix Switcher', 'RTI VIP-UHD-CTRL', 'RTI VMS-741 Multiviewer', 'RTI VMX Modular Video Distribution System', 'RTI VSS-51 (VSW-51 & VSW-51R) Scaling Switcher', 'RTI VSW-41 HDMI 2.0 Switcher', 'RTI VWS-21T Wall Plate and Scaler', 'RTI VX44 / VX88', 'RTI VXP-82', 'SnapAV B-100/300/500 Series HDMatrix Switchers', 'Vanco Evolution', 'Vaux Lattis Matrix', 'Vaux Lattis Matrix V10', 'Video Storm CMX Series', 'Visionary Solutions PacketAV 4k', 'WyreStorm HDMI Matrix', 'Zektor Matrix', 'Zigen HXL-44 Plus', 'Zigen HXL-88 Plus', 'Zigen HXL Series', 'Zigen Matrix', 'Zigen SW-41 Plus', 'Zigen SW-42 Plus', 'Amazon Fire TV', 'Android IP Remote', 'Android TV', 'Blackmagic Hyperdeck Series', 'Bose SoundTouch', 'Cicero Recording Software', 'Denon HEOS (Legacy)', 'Denon iPod Dock', 'DF Solutions', 'DLO HomeDock Pro', 'Escient Media Server', 'Fusion Movie Servers', 'Imerge / XiVA Media Server', 'Integra iPod/AirPlay Dock', 'Kaleidescape Media Server', 'KEF W1 LS50W', 'KEF W2 Media', 'KEF XIO Soundbar Media', 'Kodi Frodo/Gotham/Helix/Krypton', 'Logitech Squeezebox', 'Marantz iPod Dock', 'Marantz Network Audio Player', 'Meridian Ellipse', 'Naim Zone', 'Nuvo Wireless Audio System', 'Ovation Music Streamers', 'Pioneer Elite Network Media Player', 'Popcorn Hour C200', 'Prima Cinema Player', 'Roku Media Streamer', 'Roku Pro', 'RTiDock iPod Dock', 'Russound MBX', 'Sky Pro (UK)', 'SkyQ IP', 'Sonance iPort', 'Sonos', 'Sonos (Legacy)', 'VSSL MX', 'WD TV Live', 'WiiM', 'ADA Suite 8/16/8000 Series', 'Audac R2 Multi-zone Digital Audio Matrix', 'AudioControl Hyperion', 'AudioControl Hyperion Axis', 'AudioControl Hyperion (Legacy)', 'AVProEdge AC AEX RC HUB', 'AVProEdge AC MAX 24', 'AVPro Edge MXNet_10G (Legacy)', 'AVPro Edge MXNet_1G (Legacy)', 'AVPro Edge MXnet Dante CBOX', 'AVPro Edge MXnet Universal', 'AVPro Edge MXNet USP (Legacy)', 'B&K CT Series', 'BluOS / Bluesound / NAD Media Players', 'Blustream MV41', 'Denon HEOS', 'Elan S86AV', 'Episode Response', 'Klipsch KDA/Jamo JDA', 'Lexicon DD-8+', 'McIntosh Preamplifier', 'Meridian 218 Zone Controller', 'Monitor Audio IA Series Amplifiers', 'NuVo Grand Concerto', 'Nuvo Simplese', 'Rotel RKB Eight Channel Power Amplifier', 'RTI AD-64', 'RTI Miravue VIP-1', 'RTI Music', 'Russound E & C Series', 'Russound RNet Series', 'Sonance C4630 / DAB-1', 'Sonance DSP 8-130', 'DALCNET DGM02-1248', 'Netgear M4250 Pro AV Switches', 'Netgear M4300 Pro AV Switches', 'Unifi Network', 'Big Dog Power', 'eGauge', 'GUDE EPC Series', 'Juice Goose IP Series Power Conditioners', 'Luxul Power Distribution', 'LynTec RPC', 'Middle Atlantic Racklink', 'Netio Pro', 'Panamax Other Models', 'Panamax VT1512', 'Power Tracker', 'SurgeX Axess', 'SurgeX Squid', 'TED 5000', 'TPLink', 'TP-Link Kasa Plug', 'WattBox', 'WattBox 150/250/800', 'AJAX Jeweller Series', 'DSC Power Series', 'Elk E27', 'ELK/Ness M1', 'ELK / Ness Products', 'GE Concord', 'HAI OmniPro II', 'Hikvision DS-K2604', 'Honeywell 4232CBM', 'Honeywell EnvisaLink', 'Honeywell Vista Security Panel', 'Inner Range Inception', 'Lilin NVR', 'Lilin PTZ Cameras', 'Unifi Protect', 'Dahua DVR', 'ICRealtime DVR', 'Illumivue Surveillance', 'Lux NVR', 'Visca PTZ Cameras', 'Wirepath DVR', 'Wirepath PTZ IP Camera', 'Asterisk', 'Barco Clickshare', 'Cisco Phone Menu', 'Cisco Telepresence', 'Elmo PX Series Driver', 'LifeSize', 'LifeSize Icon', 'Polycom HDX', 'PTZOptics PT12X-SDI', 'RTI UC-IVB-50', 'RTI VXU-SC Soft Codec Video Conferencing System', 'Vaddio ClearVIEW', 'Vaddio IntelliSHOT', 'Vaddio Roboshot', 'RTI Virtual Alarm', 'RTI Virtual AV Receiver', 'RTI Virtual Blu-ray', 'RTI Virtual Climate', 'RTI Virtual HDMI Matrix', 'RTI Virtual Lighting', 'RTI Virtual Multiroom Amp', 'RTI Virtual Music Streamer', 'RTI Virtual Satellite Radio', 'RTI Virtual Satellite Receiver', 'RTI Virtual Shades', 'RTI Virtual Soundbar', 'RTI Virtual TV', 'Contemporary Research 232-ATSC+', 'DirecTV', 'Dish Network', 'Marantz Tuner ST7001', 'NuVo NV-T2SIR Tuner', 'Parasound Ztuner', 'Sirius Tuner', 'Calculator Utility', 'Clock Driver', 'Email / SMTP Client', 'Generic TCP Server', 'Global Cache', 'Global Cache Model', 'Home Butler', 'HTTP Commander', 'HTTP Server Driver', 'Hunter Douglas PowerView Gen3', 'HyperTerminal', 'ID11 Addon Driver', 'Integer Helper', 'Intrinsic Multicast Driver', 'iPort Surface Mount with Buttons', 'Kramer 2000', 'Local Flags', 'Logic Helper', 'Memory Watcher', 'One Touch Button', 'Prayer Time', 'Print To Log', 'Processor to Processor Communication', 'Pronestor Display', 'Pushover', 'QNAP Management', 'ReCalc', 'RTI System Macro Triggers', 'RTI XP Diagnostics', 'Schedule Manager', 'Shabbat Utility', 'Shelly Sense', 'Simple Logic', 'System To System', 'Text Helper', 'Two-way Strings', 'Two Way Strings', 'Virtual Presets', 'Voice Viper', 'WeMo', 'WOL Magic Packet', 'YJ-TMSv3 Tank Monitor System', 'Zigbee Network Test', 'Barco Pulse Platform (Legacy)', 'Digital Projection', 'Epson PowerLite', 'Epson Q-Series', 'Epson Universal Projector', 'Hitachi CP Series Projector', 'JVC D-ILA Projector', 'JVC LX Series DLP', 'Optoma Universal Projector', 'PJLink', 'PJLink', 'Screen Research MCS', 'Sim2 Nero 4', 'Sony 4k Projector', 'Sony VPL Video Projectors', 'V-HOME Smart Projector']

b = [items["name"] for items in a]
unique= list(set(dri_names))
counts = Counter(dri_names)
print((counts))
# for driver, count in counts.items():
#     if(count > 1):
#         print(f"{driver} (repeated {count} times)")







