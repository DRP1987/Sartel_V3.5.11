"""DM1 (Active Diagnostic Trouble Codes) decoder for J1939 CAN bus."""

from typing import Dict, List, Any, Tuple

# J1939 PGN for DM1 — Active Diagnostic Trouble Codes
DM1_PGN = 0xFECA

# Valid 2-bit lamp/flash field values (per J1939 DM1 lamp encoding used here):
#   0b00 (0) = Off / Flash off  → "Off"
#   0b01 (1) = On  / Flash on   → "On"
#   0b10 (2) = ignore           → treated as Off
#   0b11 (3) = ignore (all FFs) → treated as Off
_LAMP_ON_VALUE = 0b01

# J1939 Failure Mode Identifier descriptions
FMI_DESCRIPTIONS: Dict[int, str] = {
    0:  "Data valid but above normal operational range — most severe",
    1:  "Data valid but below normal operational range — most severe",
    2:  "Data erratic, intermittent or incorrect",
    3:  "Voltage above normal, or shorted high",
    4:  "Voltage below normal, or shorted low",
    5:  "Current below normal or open circuit",
    6:  "Current above normal or grounded circuit",
    7:  "Mechanical system not responding properly",
    8:  "Abnormal frequency, pulse width or period",
    9:  "Abnormal update rate",
    10: "Abnormal rate of change",
    11: "Root cause not known",
    12: "Bad intelligent device or component",
    13: "Out of calibration",
    14: "Special instructions",
    15: "Data valid but above normal operating range — least severe",
    16: "Data valid but above normal operating range — moderately severe",
    17: "Data valid but below normal operating range — least severe",
    18: "Data valid but below normal operating range — moderately severe",
    19: "Received network data in error",
    31: "Condition exists",
}


# ---------------------------------------------------------------------------
# SPN + FMI → human-readable error description lookup table.
#
# Key   : (spn: int, fmi: int)
# Value : str  — short error description shown to the operator
#
# Add new entries here by following the same pattern.  Any SPN/FMI pair that
# is not listed will automatically be displayed as "Unknown".
# ---------------------------------------------------------------------------
SPN_FMI_DESCRIPTIONS: Dict[Tuple[int, int], str] = {
    # ── Example entries – replace / extend these with your actual fault codes ──

    # SPN 70 – Park Brake Status
    (70, 2): "Park Brake Status Not Plausible (Vehicle Moving)",
    (70, 13): "J1939 Park Brake Switch Signal from Source CCVS1, CCVS2 or CCVS3 is missing or not available = SNA (signal not available)",
    (70, 19): "J1939 Park Brake Switch Signal from Source CCVS1, CCVS2 or CCVS3 is erractic = undefined value but not SNA",

    # SPN 70 – Park Brake Status
    (84, 0): "Vehicle Speed above programmable Threshold #1. This is not a system failure/fault.",
    (84, 2): "Vehicle Speed Signal lost",
    (84, 3): "Vehicle Speed Sensor Circuit shorted to Ubat",
    (84, 4): "Vehicle Speed Sensor Circuit shorted to GND",
    (84, 5): "Vehicle Speed Sensor open Circuit (broken wire, terminal floating)",
    (84, 6): "Vehicle Speed Sensor Anti-Tamper Detection via ABS Vehicle Speed Comparison (ABS speed and Vehicle Speed Sensor are not consistent)",
    (84, 7): "Hall effect Vehicle Speed Sensor wiring mismatch, rationality fault",
    (84, 11): "Vehicle Speed above programmable Threshold #2. This is not a system failure/fault.",
    (84, 13): "J1939 Wheel-Based Vehicle Speed Signal from Source CCVS1, CCVS2 or CCVS3 is missing or not available = SNA (signal not available)",
    (84, 14): "Hall effect Vehicle Speed Sensor supply voltage out of range",
    (84, 19): "J1939 Wheel-Based Vehicle Speed Signal from Source CCVS1, CCVS2 or CCVS3 is erractic = undefined value but not SNA",
    (84, 20): "Vehicle Speed Sensor Drifted High Error (VSS signal not plausible)",
    (84, 21): "Vehicle Speed failure (VSS Signal Not Plausible)",

    # SPN 91 – Acc pedal
    (91, 0): "Accelerator Pedal Circuit shorted to Ubat",
    (91, 2): "Accelerator Pedal out of adjustment (Learn error)",
    (91, 4): "Accelerator Pedal Circuit shorted to GND",
    (91, 7): "2-Channel Accelerator Pedal Idle Not Recognized (idle area not evaluated)",
    (91, 8): "2-Channel Accelerator Pedal Signal 1 missing",
    (91, 10): "Throttle pedal rationality check failed",
    (91, 13): "J1939 EEC2 message is missing or not available",
    (91, 14): "2-Channel Accelerator Pedal Not Learned",
    (91, 31): "2-Channel Accelerator Pedal Learned Range to Large",


    # SPN 96 – Fuel level
    (96, 13): "Fuel Level Sensor Missing or Not Available",
    (96, 19): "Fuel Level Sensor Received Network Data in Error",
    

    # SPN 107 – Air filter
    (107, 0): "Air Filter Restriction High",
    (107, 2): "Air Filter Sensor plausibility error ",
    (107, 3): "Air Filter diff.-pressure Sensor or switch shorted to Ubat",
    (107, 4): "Air Filter diff.-pressure Sensor or switch shorted to GND",

    # SPN 111 – Coolant level
    (111, 1): "Coolant Level below safe operating level – (SEL Condition)",
    (111, 3): "Coolant Level Circuit shorted to Ubat",
    (111, 4): "Coolant Level Sensor Circuit Failed Low",
    (111, 6): "Coolant Level Sensor Circuit shorted to GND",
    (111, 18): "Coolant Level below operating level (pre-warning condition)",

    # SPN 120 – RF
    (120, 13): "J1939 Retarder Fluid Message is missing or not available (J1939 Cabin Message is missing or not available?)",
    
    # SPN 158 – KL15
    (158, 2): "KL15 ignition switch status of CPC and MCM do not match.",

    # SPN 168 – Battery
    (168, 0): "Battery voltage to high",
    (168, 1): "Battery voltage very Low",
    (168, 7): "Opt Idle Detected Charging System or Battery Failure",
    (168, 9): "Main battery connection lost",
    (168, 14): "ECU powerdown not completed (Main Battery Terminal Possibly Floating)",
    (168, 18): "Battery voltage low",

    # SPN 171 – Amb temperature
    (171, 2): "Ambient temperature sensor data Erratic",
    (171, 3): "Ambient temperature Sensor shortted to UBat",
    (171, 4): "Ambient temperature sensor shortted to GND",
    (171, 9): "J1587 Ambient Air Temp Sensor Data Message Stopped Arriving",
    (171, 14): "J1587 Ambient Air Temp Sensor Data Not Received This Ign Cycle",

    # SPN 187 – FUSO volume
    (187, 3): "Idle Volume Sensor circuit shorted to Ubat",
    (187, 4): "Idle Volume Sensor circuit shorted to GND",

    # SPN 191 – output speed shaft
    (191, 9): "J1939 ETC1 Message is missing or not available",
    (191, 13): "J1939 Transmission Output Shaft Speed Signal is missing or not available = SNA (signal not available)",
    (191, 19): "J1939 Transmission Output Shaft Speed Signal is missing or not available = SNA (signal not available)",

    # SPN 247 – Hours/ Mileage MCM&ACM
    (247, 0): "MCM Engine Hours Data higher than expected",
    (247, 1): "MCM Engine Hours Data lower than expected",
    (247, 9): "MCM Engine Hours Data not received or stopped arriving",
    (247, 10): "MCM Engine Hours Data not received or stopped arriving",
    (247, 14): "ACM Reported Ash Mileage is Lower then the CPC Stored Value",

    # SPN 523 – Transmission Gear 
    (523, 13): "J1939 Transmission Current Gear Signal is missing or Not available",
    (523, 19): "J1939 Transmission Current Gear Signal is erractic = undefined value but not SNA",

     # SPN 524 – ETC2 message
    (524, 9): "J1939 ETC2 message is missing or not available",

    # SPN 525 – Transmission Gear
    (525, 7): "Transmission gear selection switch reports internal error.",
    (525, 9): "J1939 Powertrain Message (transfer case / PTO) is missing",
    (525, 19): "Transmission gear selection switch reports unplausible engine brake stage requests.",

    # SPN 527 – CCVS
    (527, 9): "J1939 CCVS is missing or not available",

    # SPN 556 – RC
    (556, 9): "J1939 RC Message from Transmission Retarder is missing",

    # SPN 558 – Idle switch
    (558, 2): "Idle Validation Switch Inputs Reversed.",
    (558, 3): "Idle Validation Switch 1 Circuit shorted to Ubat. The two idle switches are not synchron (check AP)",
    (558, 4): "Idle Validation Switch 1 Circuit shorted to GND. The two idle switches are not synchron (check AP)",
    (558, 5): "Idle Validation Switch 2 Circuit shorted to GND. The two idle switches are not synchron (check AP)",
    (558, 6): "Idle Validation Switch 2 Circuit shorted to Ubat. The two idle switches are not synchron (check AP)",

    # SPN 569 – EAC1
    (569, 9): "J1939 EAC1 Message is missing or not available",

    # SPN 571 – Engine brake
    (571, 4): "Engine Brake Disable push-button shorted to Ground or pressed too long",

    # SPN 596 – Cruise control
    (596, 13): "J1939 Cruise Control Enable Switch Signal from Source CCVS1, CCVS2 or CCVS3 missing or not available = SNA (signal not available)",
    (596, 19): "J1939 Cruise Control Enable Switch Signal from Source CCVS1, CCVS2 or CCVS3 erractic = undefined value but not SNA",

     # SPN 597 – Service brake
    (597, 2): "Idle Validation Switch Inputs Reversed.",
    (597, 13): "Idle Validation Switch 1 Circuit shorted to Ubat. The two idle switches are not synchron (check AP)",
    (597, 19): "J1939 Service Brake Switch Signal from Source CCVS1, CCVS2 or CCVS3 erractic = undefined value but not SNA",

    # SPN 598 – Clutch switch
    (598, 2): "Clutch switch status not plausible",

    # SPN 599 – Cruise control
    (599, 4): "Cruise Control SET and RESUME Circuits shorted to GND (SET and RESUME applied at the same time)",

    # SPN 600 – Cruise control
    (600, 13): "J1939 Cruise Control Coast Switch Signal from Source CCVS1, CCVS2 or CCVS3 missing or not available = SNA (signal not available)",
    (600, 19): "J1939 Cruise Control Coast Switch Signal from Source CCVS1, CCVS2 or CCVS3 erractic = undefined value but not SNA",


    # SPN 3490 – DPF Purge lamp
    (3490, 3): "DEF Purge Lamp Circuit Failed High - Shortcut to batt fault",
    (3490, 4): "DEF Purge Lamp Circuit Failed High - Shortcut to ground fault",
    (3490, 5): "DEF Purge Lamp Circuit Failed High - Detect open load fault",

    # SPN 110 – Engine Coolant Temperature
    (110, 0): "Engine coolant temperature above normal range",
    (110, 3): "Engine coolant temperature sensor voltage high",
    (110, 4): "Engine coolant temperature sensor voltage low",

    # SPN 190 – Engine Speed
    (190, 0): "Engine speed above normal operating range",
    (190, 8): "Engine speed signal abnormal frequency",

    # SPN 91 – Accelerator Pedal Position
    (91, 2):  "Accelerator pedal signal erratic or intermittent",
    (91, 3):  "Accelerator pedal sensor voltage high",
    (91, 4):  "Accelerator pedal sensor voltage low",

    # SPN 1569 – Engine Protection Torque Derate
    (1569, 31): "Engine protection derate active – condition exists",
}


def get_dtc_description(spn: int, fmi: int) -> str:
    """Return the error description for a given SPN/FMI pair, or 'Unknown'."""
    return SPN_FMI_DESCRIPTIONS.get((spn, fmi), "Unknown")


def _lamp_on(byte: int, shift: int) -> bool:
    """Return True only if the 2-bit field at *shift* is exactly 0b01 (On/Flash-on).

    Any value other than 0b01 is treated as Off/ignore per the DM1 spec used here.
    This prevents 0xFF bytes from spuriously activating all lamps.
    """
    return ((byte >> shift) & 0x03) == _LAMP_ON_VALUE


def decode_dm1(data: List[int]) -> Dict[str, Any]:
    """
    Decode a DM1 (Active DTCs) J1939 message from PGN 0xFECA.

    Message byte layout (0-indexed, little-endian SPNs):
    • Byte 0  — Lamp on/off status
                  bits 1-0 : Check Engine Lamp      (0b01=On, others=Off/ignore)
                  bits 3-2 : Amber Warning Lamp      (0b01=On, others=Off/ignore)
                  bits 5-4 : Red Stop Lamp           (0b01=On, others=Off/ignore)
                  bits 7-6 : Aftertreatment Lamp     (0b01=On, others=Off/ignore)
                  NOTE: 0xFF → all 2-bit fields are 0b11 → all lamps Off (ignored)
    • Byte 1  — Flash lamp status (same bit layout as byte 0)
                  bits 1-0 : Flash Check Engine Lamp (0b01=Flash, others=no-flash/ignore)
                  bits 3-2 : Flash Amber Warning Lamp
                  bits 5-4 : Flash Red Stop Lamp
                  bits 7-6 : Flash Aftertreatment Lamp
                  Lamp status = "Flash" when byte-0 bit = On AND byte-1 bit = Flash,
                                "On"    when byte-0 bit = On AND byte-1 bit ≠ Flash,
                                "Off"   otherwise.
    • First DTC (bytes 2–4):
                  bytes 2-3: SPN bits  0-15 (little-endian)
                  byte  4  : bits 7-5 = SPN bits 18-16 ; bits 4-0 = FMI
    • Each additional DTC (5 bytes each, starting at byte 5):
                  bytes +0,+1 : lamp status (same layout as byte 0)
                  bytes +2,+3 : SPN bits 0-15
                  byte  +4   : SPN bits 18-16 (bits 7-5) + FMI (bits 4-0)

    Args:
        data: Raw CAN message bytes as a list of ints.

    Returns:
        Dict with keys:
        • 'lamps'  — dict of lamp names → status strings ("Off", "On", "Flash")
        • 'dtcs'   — list of dicts, each with 'spn', 'fmi', 'fmi_desc'
    """
    result: Dict[str, Any] = {
        'lamps': {
            'check_engine':    'Off',
            'amber_warning':   'Off',
            'red_stop':        'Off',
            'aftertreatment':  'Off',
        },
        'dtcs': [],
    }

    if not data:
        return result

    # --- Lamp on/off status (byte 0) + flash status (byte 1) ---
    lamp_byte  = data[0]
    flash_byte = data[1] if len(data) > 1 else 0x00

    # Bit layout: bits 0-1 = check_engine, 2-3 = amber_warning,
    #             4-5 = red_stop,           6-7 = aftertreatment
    for key, shift in (
        ('check_engine',   0),
        ('amber_warning',  2),
        ('red_stop',       4),
        ('aftertreatment', 6),
    ):
        is_on    = _lamp_on(lamp_byte,  shift)
        is_flash = _lamp_on(flash_byte, shift)
        if is_on and is_flash:
            result['lamps'][key] = 'Flash'
        elif is_on:
            result['lamps'][key] = 'On'
        # else remains 'Off'

    # --- First DTC at bytes 2-4 (need at least 5 bytes total) ---
    idx = 2
    if len(data) >= idx + 3:
        spn_byte0 = data[idx]
        spn_byte1 = data[idx + 1]
        fmi_byte  = data[idx + 2]
        # Skip DTC entry when all three bytes are 0x00 (no DTC data present)
        if spn_byte0 != 0x00 or spn_byte1 != 0x00 or fmi_byte != 0x00:
            spn = spn_byte0 | (spn_byte1 << 8) | ((fmi_byte >> 5) << 16)
            fmi = fmi_byte & 0x1F
            result['dtcs'].append({
                'spn':        spn,
                'fmi':        fmi,
                'fmi_desc':   FMI_DESCRIPTIONS.get(fmi, f'FMI {fmi}'),
                'error_desc': get_dtc_description(spn, fmi),
            })
        idx += 3  # advance past first DTC

    # --- Additional DTCs: 5-byte groups [2 lamp bytes, 2 SPN bytes, 1 FMI byte] ---
    while idx + 5 <= len(data):  # need bytes idx … idx+4 inclusive (5 bytes total)
        # Bytes idx and idx+1 are per-DTC lamp status — skip them
        spn_byte0 = data[idx + 2]
        spn_byte1 = data[idx + 3]
        fmi_byte  = data[idx + 4]
        # Skip DTC entry when all three bytes are 0x00 (no DTC data present)
        if spn_byte0 != 0x00 or spn_byte1 != 0x00 or fmi_byte != 0x00:
            spn = spn_byte0 | (spn_byte1 << 8) | ((fmi_byte >> 5) << 16)
            fmi = fmi_byte & 0x1F
            result['dtcs'].append({
                'spn':        spn,
                'fmi':        fmi,
                'fmi_desc':   FMI_DESCRIPTIONS.get(fmi, f'FMI {fmi}'),
                'error_desc': get_dtc_description(spn, fmi),
            })
        idx += 5

    return result
