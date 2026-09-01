"""
CSE Account Assignment Tool  —  Streamlit App
Authenticates to Salesforce via username/password/security token,
fetches two reports, applies RSE→CSE assignment algorithm, exports to Excel.
"""

import io, re, time, math, xml.etree.ElementTree as ET
from collections import defaultdict

import requests
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CSE Assignment Tool",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS  —  SAP Light theme
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Fonts & base ── */
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', '72', sans-serif; }

/* ── Top bar ── */
.tool-header {
    background: #0070F2;
    border-radius: 8px;
    padding: 20px 28px 16px 28px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 18px;
}
.tool-header h1 {
    color: #ffffff !important;
    font-size: 1.45rem;
    font-weight: 700;
    margin: 0;
    line-height: 1.2;
}
.tool-header p {
    color: #cfe6ff;
    font-size: 0.85rem;
    margin: 4px 0 0 0;
}

/* ── Metric cards ── */
.metric-card {
    background: #F5F6F7;
    border: 1px solid #E5E5E5;
    border-radius: 8px;
    padding: 18px 22px;
    text-align: center;
}
.metric-card .label {
    font-size: 0.78rem;
    color: #6a6d70;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}
.metric-card .value {
    font-size: 2rem;
    font-weight: 700;
    color: #0070F2;
    line-height: 1.1;
}
.metric-card .value.green  { color: #107E3E; }
.metric-card .value.orange { color: #E9730C; }

/* ── Info box ── */
.info-box {
    background: #E1F4FF;
    border-left: 4px solid #4CB1FF;
    border-radius: 4px;
    padding: 12px 16px;
    font-size: 0.85rem;
    color: #32363A;
    margin: 8px 0 16px 0;
}

/* ── Section headers ── */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #32363A;
    border-bottom: 2px solid #0070F2;
    padding-bottom: 6px;
    margin: 20px 0 12px 0;
}

/* ── Override tag ── */
.override-tag {
    display: inline-block;
    background: #FFF3CD;
    border: 1px solid #F0AD4E;
    border-radius: 12px;
    padding: 2px 10px;
    font-size: 0.75rem;
    font-weight: 600;
    color: #856404;
    margin-left: 6px;
}

/* ── Hide Streamlit default footer ── */
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
SF_INSTANCE      = "https://sapconcur.my.salesforce.com"
SF_API_VER       = "v59.0"
RPT_NEW_ACCOUNTS = "00O0e000005i4gg"
RPT_VOLUMES      = "00O7V000006IT4i"
MIN_ACCOUNTS     = 20

# ─────────────────────────────────────────────────────────────────────────────
# NAME NORMALIZATION
# ─────────────────────────────────────────────────────────────────────────────
_NORM_MAP = {
    "Joe Bellefeuille":   "Joseph Bellefeuille",
    "Nathaniel Heussner": "Nate Heussner",
    "Stephen Snediker":   "Steve Snediker",
    "Tom Wahl":           "Thomas Wahl",
    "Joe Silva":          "Joseph Silva",
    "Chris Spencer":      "Christopher Spencer",
    "Rachel Schmidt":     "Rachel Meyer",
    "Mike Monello":       "Michael Monello",
}

def norm(name: str) -> str:
    if not name:
        return ""
    s = str(name).strip()
    s = re.sub(r'\s*\(.*?\)\s*$', '', s).strip()
    return _NORM_MAP.get(s, s)

def norm_csm(raw: str) -> str:
    s = (raw or "").strip()
    if "@" in s:
        local = s.split("@")[0]
        parts = re.split(r'[._]', local)
        s = " ".join(p.capitalize() for p in parts if p)
    return norm(s)

NO_PARTNERSHIP_CSMS = {"Client Success Management", "Round Robin CSM", "Round Robin", ""}

# ─────────────────────────────────────────────────────────────────────────────
# EMBEDDED CSE ROSTER  (source: CSEs.xlsx — Kylie Barrett excluded)
# ─────────────────────────────────────────────────────────────────────────────
_RAW_CSES = [
    # (salesforce_id, raw_name, manager, region, segment)
    # ── Key ──────────────────────────────────────────────────────────────────
    ("0050e0000084hShAAI", "Jon Delwiche",         "Andrew Hop",    "US SMB CSE Key Central",       "Key"),
    ("00532000005GKmYAAW", "Tony Pacheco",          "Peter Gadd",    "US SMB CSE Key Central",       "Key"),
    ("0050e000007OlveAAC", "Leah Harris",           "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("00532000005R9nkAAC", "Randi Kruger",          "Peter Gadd",    "US SMB CSE Key South",         "Key"),
    ("005Pg00000SDPrhIAH", "Jacob Nickoloff",       "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("00560000003M0SfAAK", "Anders Halvorsen",      "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("0057V00000BfS89QAF", "Jonathan Barth",        "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0057V000008UkHZQA0", "Natalie Rizk",          "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0057V000009JfmFQAS", "Tyler Klein",           "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("0057V000009JrneQAC", "Oliver Holdenson",      "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("005Pg00000Rx6mbIAB", "Libby Hartnagel",       "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("005Pg00000P5ttVIAR", "Austin Aghamirzai",     "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("005Pg00000PMiqPIAT", "Teylen Sheesley",       "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0057V00000BfSJqQAN", "Christopher Spencer",   "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("0050e000006YpbZAAS", "Ryan Doyle",            "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0050e000007olKpAAI", "Mike Antkowiak",        "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0057V000008UgfXQAS", "Alden Martinez",        "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("005Pg00000PDl2jIAD", "Brianna Basolo",        "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("0050e000006iRHfAAM", "Joe Bellefeuille",      "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("0057V000009fXzjQAE", "Joe Vigil",             "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("0057V000008iHOUQA2", "Mark Hemmerle",         "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("0050e0000077RXJAA2", "Tyler Krob",            "Marissa Mock",  "US SMB CSE Key Great Lakes",   "Key"),
    ("0057V000009Ju27QAC", "Brooke Mullis",         "Ronit Cohn",    "US SMB CSE Key North",         "Key"),
    ("0057V000009cmzYQAQ", "Michaela Gormley",      "Heather Lewis", "US SMB CSE Key West",          "Key"),
    ("005Pg00000DwTzqIAF", "David Jensen",          "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("0057V00000B57XbQAJ", "Jack Zabel",            "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("0057V000009JCz6QAG", "Tyler Sanford",         "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("0057V000008RgeiQAC", "Rachel Meyer",          "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("0050e000006eyPUAAY", "Ben Angelo",            "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    ("0050e0000070PpvAAE", "Thang Nguyen",          "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    ("0050e000006x5stAAA", "Jordan Breisacher",     "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    ("0057V00000B57YZQAZ", "Scott Bere",            "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("005Pg00000FunJxIAJ", "Karrah Manzanarez",     "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    ("0057V000008UgkrQAC", "Tom Larson",            "Randi Kruger",  "US SMB CSE Key South",         "Key"),
    ("005Pg00000QSHi6IAH", "Mackenzie Bowen",       "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    ("005Pg00000Nbz45IAB", "Kelsey Fredrickson",    "Jake Rutenbar", "US SMB CSE Key Mid Atlantic",  "Key"),
    # ── Strategic ────────────────────────────────────────────────────────────
    ("0050e000007ovW7AAI", "Zack Smith",            "Brooke Nelson",    "US SMB CSE Strategic Northwest", "Strategic"),
    ("0057V000008iA3UQAU", "Dustin Misener",        "Nicole Carlsen",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("0057V000009JD6BQAW", "Brittani Goins",        "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0050e000006Ky3YAAS", "Murphy Donovan",        "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("00532000005yf9YAAQ", "Christian Larson",      "Andrew Hop",       "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0057V000008iME0QAM", "Lindsay Paxton",        "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("0050e000007oSbDAAU", "Samantha O'Connell",    "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("00532000005EzQmAAK", "Joel Segall",           "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("00532000005rs5dAAA", "Ty Saathoff",           "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("0057V000009KSbQQAW", "Aaron Korus",           "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0057V000008VYA4QAO", "Tom Wahl",              "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("0050e000006qJSqAAM", "Lauren Pellowski",      "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0050e0000078FprAAE", "Zack Scharf",           "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0057V000008RgcIQAS", "Colin Kraker",          "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0057V000008W9eLQAS", "Andrea Flor",           "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("00532000005G3IZAA0", "Kevin Chheang",         "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("0057V000008QVr5QAG", "Joseph Zangel",         "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0057V000008VrymQAC", "Joe Dorey",             "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("0057V000008XS1FQAW", "Mara Obermeier",        "Amanda Meek",      "US SMB CSE Strategic Mountain",  "Strategic"),
    ("0057V000009JfmPQAS", "Jill Desjardine",       "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("005Pg00000UwEm5IAF", "Nicole Breitenstein",   "Dave Elinger",     "US SMB CSE Strategic Southwest", "Strategic"),
    ("0050e000006YpchAAC", "Joseph Silva",          "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("005Pg00000IH5fFIAT", "Evan Smith",            "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("0050e000007oSffAAE", "Laura Jungbauer",       "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0050e000007qwOhAAI", "Matt Knight",           "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("0050e000006x9OdAAI", "Tyler Hazen",           "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("0057V000009QX2DQAW", "Mark Workman",          "Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("00532000005sE95AAE", "Nathaniel Heussner",    "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("005Pg00000ObrNhIAJ", "Michael Monello",       "Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("0057V000008QGIjQAO", "Debbie Saysanavongphet","Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("00532000006DMbJAAW", "Brittany Whims",        "Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("005Pg00000FC7ZlIAL", "Tyler Nuquay",          "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("00532000005GsOxAAK", "Evan Anderson",         "Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("00560000003e4voAAA", "Rachel Burns",          "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0050e000007RAOeAAO", "Jeffrey Danner",        "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0050e000008gXTWAA2", "Jessica Klein",         "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("0050e000006YxyCAAS", "Jun Park",              "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("0050e000008hkYDAAY", "Dan Eagen",             "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0050e000006yz1nAAA", "Emily Norris",          "Megan Frodge",     "US SMB CSE Strategic Northeast", "Strategic"),
    ("005320000050NAzAAM", "Jamie Stewart",         "Blake Karnes",     "US SMB CSE Strategic Northwest", "Strategic"),
    ("0057V000008iGMDQA2", "Anna Christofaro",      "Christian Larson", "US SMB CSE Strategic Midwest",   "Strategic"),
    ("0050e0000078Z5PAAU", "Jordan Buri",           "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    ("0057V00000ASympQAD", "Trevor Hecht",          "Samantha Young",   "US SMB CSE Strategic Southeast", "Strategic"),
    # ── Premier ──────────────────────────────────────────────────────────────
    ("00532000005CmTZAA0", "Andy Brinkhaus",   "Amanda Meek",   "US SMB CSE Premier Midwest",   "Premier"),
    ("00532000006D0kEAAS", "Kyle Loving",      "Jason Rainey",  "US SMB CSE Premier East",      "Premier"),
    ("00560000001POCcAAO", "Mark Garretson",   "Caryn Ross",    "US SMB CSE Premier Mountain",  "Premier"),
    ("0050e000007dtmVAAQ", "Zak Priest",       "",              "US SMB CSE Premier North",     "Premier"),
    ("0050e0000078qPhAAI", "Cassie Mckinley",  "Caryn Ross",    "US SMB CSE Premier Mountain",  "Premier"),
    ("00532000005fQkvAAE", "Rob Meek",         "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("0057V000008XS0RQAW", "Ashley McCue",     "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("0050e000007Qb2xAAC", "Adam Sala",        "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("0057V000008WtN1QAK", "Alanna Parrott",   "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("00560000002TJb2AAG", "Stephen Snediker", "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("0050e000006AXdtAAG", "Julie Barter",     "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("0057V000009K19AQAS", "Robert Stoering",  "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("005600000044z0zAAA", "Katie Byrnes",     "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("0050e000006Z2BtAAK", "Ben Goman",        "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("0050e000007b5MdAAI", "Alex Capeloto",    "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("00532000005DWEGAA4", "Michael Landes",   "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("00532000005OACmAAO", "Beth Danielson",   "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("005600000046mvDAAQ", "Mandy Gallanar",   "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("0057V000008hq3yQAA", "Lindsay Wilson",   "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("00532000005LW0UAAW", "Kim Koehn",        "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("005600000046txpAAA", "Deb Tegan",        "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("00532000005EpKDAA0", "Katrina Brock",    "Chris Smith",   "US SMB CSE Premier North",     "Premier"),
    ("00532000005EQ6IAAW", "Sam Angelo",       "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("005320000061QpuAAE", "Amy Lawrence",     "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("0057V000009ImN5QAK", "Sabah Ozan",       "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("0050e000007qoZwAAI", "Meaghan Rodgers",  "Brooke Nelson", "US SMB CSE Premier South",     "Premier"),
    ("0050e000006jfCWAAY", "Breck Hansen",     "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("0050e000006PkRpAAK", "Deb Smith",        "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("0053200000501hzAAA", "Danny Lewis",      "Angie Koplan",  "US SMB CSE Premier West",      "Premier"),
    ("00532000004kkFyAAI", "Natalie Jamieson", "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("0050e0000078G6nAAE", "Blair Sievert",    "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("00560000004da8wAAA", "Ellen Bastian",    "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
    ("0050e000007oSwLAAU", "Aghiles Benali",   "Kyle Loving",   "US SMB CSE Premier East",      "Premier"),
]

CSE_BY_NORM = {}
for _sid, _raw, _mgr, _reg, _seg in _RAW_CSES:
    _n = norm(_raw)
    CSE_BY_NORM[_n] = {"id": _sid, "name": _raw, "norm": _n,
                       "manager": _mgr, "region": _reg, "segment": _seg}

# ─────────────────────────────────────────────────────────────────────────────
# EMBEDDED PARTNERSHIPS
# ─────────────────────────────────────────────────────────────────────────────
_STRAT_P = {
    "Alex Schlatter":     ["Evan Smith", "Jordan Buri", "Mark Workman"],
    "Alison Soukthavone": ["Brittany Whims", "Anna Christofaro", "Jessica Klein", "Kevin Chheang"],
    "Ian McGill":         ["Tyler Nuquay", "Debbie Saysanavongphet", "Evan Anderson", "Jamie Stewart", "Kaitlin Dailey"],
    "Jamie Holub":        ["Colin Kraker", "Joel Segall", "Joseph Zangel", "Lauren Pellowski", "Zack Scharf"],
    "Josh Shoun":         ["Andrea Flor", "Ashley McCue", "Joe Silva", "Mara Obermeier", "Trevor Hecht", "Ty Saathoff"],
    "Melissa Yapp":       ["Jeffrey Danner", "Laura Jungbauer", "Mike Monello"],
    "Mitch Pehrson":      ["Nate Heussner", "Tyler Hazen", "Joe Dorey"],
    "Nicole Ramtahal":    ["Emily Norris", "Jun Park", "Lindsay Paxton", "Matt Knight"],
    "Taylor Augustin":    ["Dan Eagen", "Aaron Korus", "Samantha O'Connell", "Tom Wahl"],
}
_PREMIER_P = {
    "Alex Schlatter":     ["Katie Byrnes", "Kim Koehn", "Michael Landes"],
    "Alison Soukthavone": ["Amy Lawrence", "Katrina Brock", "Julie Barter"],
    "Ian McGill":         ["Alex Capeloto", "Jon Salmon", "Danny Lewis"],
    "Jamie Holub":        ["Aghiles Benali", "Deb Smith", "Sam Angelo"],
    "Josh Shoun":         ["Adam Sala", "Ben Goman", "Rob Meek"],
    "Melissa Yapp":       ["Lindsay Wilson", "Robert Stoering", "Stephen Snediker"],
    "Mitch Pehrson":      ["Mandy Gallanar", "Meaghan Rodgers", "Sabah Ozan"],
    "Nicole Ramtahal":    ["Alanna Parrott", "Breck Hansen", "Ellen Bastian"],
    "Taylor Augustin":    ["Blair Sievert", "Deb Tegan", "Natalie Jamieson"],
}
_KEY_P = {
    "Grace Kremer":  ["Alden Martinez", "Austin Aghamirzai", "Brianna Basolo", "Brooke Mullis",
                      "Chris Spencer", "Jacob Nickoloff", "Jonathan Barth", "Jordan Breisacher",
                      "Karrah Manzanarez", "Kelsey Fredrickson", "Natalie Rizk", "Ryan Doyle",
                      "Teylen Sheesley", "Thang Nguyen", "Libby Hartnagel", "Tyler Krob"],
    "Rizza Bautista": ["Anders Halvorsen", "Mackenzie Bowen", "Ben Angelo", "David Jensen",
                       "Jack Zabel", "Joe Bellefeuille", "Joe Vigil", "Mark Hemmerle",
                       "Michaela Gormley", "Mike Antkowiak", "Oliver Holdenson", "Rachel Schmidt",
                       "Scott Bere", "Tom Larson", "Tyler Klein", "Tyler Sanford"],
}
_ALL_PARTNERSHIPS = {"Strategic": _STRAT_P, "Premier": _PREMIER_P, "Key": _KEY_P}

def get_partnered_cses(segment: str, csm_raw: str) -> list:
    csm = norm_csm(csm_raw)
    if csm in NO_PARTNERSHIP_CSMS:
        return []
    pool = _ALL_PARTNERSHIPS.get(segment, {})
    for key, cses in pool.items():
        if norm(key) == norm(csm):
            return [norm(c) for c in cses]
    return []

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE API
# ─────────────────────────────────────────────────────────────────────────────
def sf_headers(sid: str) -> dict:
    return {"Authorization": f"Bearer {sid}", "Accept": "application/json",
            "Content-Type": "application/json"}

def sf_login(email: str, password: str, security_token: str) -> tuple:
    """
    Authenticate via the Salesforce SOAP Partner API.
    Returns (session_id, instance_url).

    Unlike the browser sid cookie, the session produced here is an API
    session — not bound to the originating IP address — so it works from
    any server, including Streamlit Cloud.
    """
    soap = f"""<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope
    xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:urn="urn:partner.soap.sforce.com">
  <soapenv:Body>
    <urn:login>
      <urn:username>{email}</urn:username>
      <urn:password>{password}{security_token}</urn:password>
    </urn:login>
  </soapenv:Body>
</soapenv:Envelope>"""

    try:
        resp = requests.post(
            "https://login.salesforce.com/services/Soap/u/59.0",
            data=soap.encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=UTF-8", "SOAPAction": "login"},
            timeout=30,
        )
    except requests.exceptions.Timeout:
        raise TimeoutError("Login request timed out. Check your network connection.")
    except Exception as e:
        raise RuntimeError(f"Login request failed: {e}")

    root = ET.fromstring(resp.text)
    ns_env = "http://schemas.xmlsoap.org/soap/envelope/"
    ns_sf  = "urn:partner.soap.sforce.com"

    # SOAP faults: faultstring lives in NO namespace (bare element), not in the
    # envelope namespace — search both to be safe.
    fault = (
        root.find(".//faultstring") or
        root.find(f".//{{{ns_env}}}faultstring")
    )
    if fault is not None:
        raise PermissionError(f"Salesforce login failed: {fault.text}")

    # sessionId may be in the sf namespace or in no namespace depending on SFDC version
    session_id = (
        root.findtext(f".//{{{ns_sf}}}sessionId") or
        root.findtext(".//sessionId")
    )
    server_url = (
        root.findtext(f".//{{{ns_sf}}}serverUrl") or
        root.findtext(".//serverUrl")
    )

    if not session_id:
        # Surface the raw response so the user can see what Salesforce actually said
        raise PermissionError(
            f"Salesforce did not return a session ID (HTTP {resp.status_code}). "
            f"Raw response: {resp.text[:400]}"
        )

    # Derive the instance URL from the serverUrl
    # e.g. https://sapconcur.my.salesforce.com/services/Soap/...  →  https://sapconcur.my.salesforce.com
    instance_url = SF_INSTANCE
    if server_url:
        from urllib.parse import urlparse
        parsed = urlparse(server_url)
        instance_url = f"{parsed.scheme}://{parsed.netloc}"

    return session_id, instance_url

def parse_sf_report(data: dict) -> list:
    meta   = data.get("reportMetadata", {})
    ext    = data.get("reportExtendedMetadata", {})
    cols   = meta.get("detailColumns", [])
    ci     = ext.get("detailColumnInfo", {})
    labels = [ci.get(c, {}).get("label", c) for c in cols]
    rows_raw = data.get("factMap", {}).get("T!T", {}).get("rows", [])
    rows = []
    for r in rows_raw:
        cells = r.get("dataCells", [])
        row = {}
        for i, lbl in enumerate(labels):
            if i < len(cells):
                cell = cells[i]
                row[lbl] = cell.get("label") if cell.get("label") not in (None, "") else cell.get("value", "")
        rows.append(row)
    return rows

def test_connection(sid: str) -> dict:
    """
    Lightweight 3-step connection check — no report data is fetched.

    Step 1: GET /services/data/v59.0/          — verifies the session is valid
    Step 2: GET /services/oauth2/userinfo       — returns the logged-in user's name
    Step 3: GET /analytics/reports/{id}/describe — confirms each report is accessible
                                                   (metadata only, does NOT run the report)

    Returns:
        {
          "steps": [{"label": str, "ok": bool, "detail": str}, ...],
          "user":  str,   # display name from identity endpoint
          "ok":    bool,  # True only if all steps passed
        }
    """
    hdrs  = sf_headers(sid)
    steps = []
    user  = ""

    # ── Step 1: basic auth ─────────────────────────────────────────────────────
    try:
        r = requests.get(
            f"{SF_INSTANCE}/services/data/{SF_API_VER}/",
            headers=hdrs, timeout=15,
        )
        if r.status_code == 401:
            steps.append({"label": "Session valid", "ok": False,
                          "detail": "Session ID is invalid or expired. Copy a fresh sid cookie."})
            return {"steps": steps, "user": user, "ok": False}
        r.raise_for_status()
        steps.append({"label": "Session valid", "ok": True, "detail": "Salesforce accepted the session ID."})
    except requests.exceptions.Timeout:
        steps.append({"label": "Session valid", "ok": False,
                      "detail": "Request timed out. Salesforce is likely blocking API calls from "
                                "this server's IP address. Use CSV upload instead."})
        return {"steps": steps, "user": user, "ok": False}
    except Exception as e:
        steps.append({"label": "Session valid", "ok": False, "detail": f"Connection error: {e}"})
        return {"steps": steps, "user": user, "ok": False}

    # ── Step 2: identity / logged-in user ──────────────────────────────────────
    try:
        ur = requests.get(
            f"{SF_INSTANCE}/services/oauth2/userinfo",
            headers=hdrs, timeout=15,
        )
        if ur.ok:
            ud   = ur.json()
            user = ud.get("name") or ud.get("preferred_username") or ud.get("email") or ""
            steps.append({"label": "User identity", "ok": True,
                          "detail": f"Logged in as: {user}"})
        else:
            steps.append({"label": "User identity", "ok": False,
                          "detail": f"Could not retrieve user info (HTTP {ur.status_code})."})
    except Exception as e:
        steps.append({"label": "User identity", "ok": False, "detail": f"Identity check failed: {e}"})

    # ── Step 3: report access (describe — metadata only, no rows fetched) ──────
    for rpt_name, rpt_id in [
        ("New Accounts report", RPT_NEW_ACCOUNTS),
        ("Account Volumes report", RPT_VOLUMES),
    ]:
        try:
            rr = requests.get(
                f"{SF_INSTANCE}/services/data/{SF_API_VER}/analytics/reports/{rpt_id}/describe",
                headers=hdrs, timeout=15,
            )
            if rr.status_code == 200:
                meta  = rr.json()
                rname = meta.get("reportMetadata", {}).get("name", rpt_id)
                steps.append({"label": rpt_name, "ok": True,
                              "detail": f"Accessible — \"{rname}\""})
            elif rr.status_code == 404:
                steps.append({"label": rpt_name, "ok": False,
                              "detail": f"Report {rpt_id} not found. Check the report ID."})
            elif rr.status_code == 401:
                steps.append({"label": rpt_name, "ok": False,
                              "detail": "Access denied. Your user may not have permission to this report."})
            else:
                steps.append({"label": rpt_name, "ok": False,
                              "detail": f"HTTP {rr.status_code}: {rr.text[:120]}"})
        except requests.exceptions.Timeout:
            steps.append({"label": rpt_name, "ok": False,
                          "detail": "Request timed out. Analytics API may be blocked from this server."})
        except Exception as e:
            steps.append({"label": rpt_name, "ok": False, "detail": f"Check failed: {e}"})

    all_ok = all(s["ok"] for s in steps)
    return {"steps": steps, "user": user, "ok": all_ok}


def fetch_report(sid: str, report_id: str, status_fn=None, timeout_s: int = 90) -> list:
    """
    Fetch all rows from a Salesforce Analytics report.

    Strategy:
      1. Try the synchronous API first (?includeDetails=true).
         - For small reports (new accounts) this returns immediately with allData=true.
         - Fast, no queueing, no polling.
      2. If allData=false (>2 000 rows), fall back to the async Instances API.
         - Gives Salesforce 3 s to start the job, then polls every 4 s.
      3. If async times out, return whatever sync gave us and set a warning flag
         so the caller can surface it to the user.
    """
    base  = f"{SF_INSTANCE}/services/data/{SF_API_VER}/analytics/reports/{report_id}"
    hdrs  = sf_headers(sid)
    sync_rows = []

    # ── Step 1: synchronous fetch ──────────────────────────────────────────────
    if status_fn:
        status_fn(f"Fetching report {report_id}…")
    try:
        resp = requests.get(f"{base}?includeDetails=true", headers=hdrs, timeout=60)
        if resp.status_code == 401:
            raise PermissionError("Invalid or expired session ID.")
        resp.raise_for_status()
        data      = resp.json()
        sync_rows = parse_sf_report(data)
        if data.get("allData", True):
            return sync_rows          # ← all rows returned; done
        if status_fn:
            status_fn(f"Report has >2 000 rows — switching to async fetch…")
    except PermissionError:
        raise
    except Exception as sync_err:
        if status_fn:
            status_fn(f"Sync attempt failed ({sync_err}) — trying async…")

    # ── Step 2: async fallback ─────────────────────────────────────────────────
    try:
        r2 = requests.post(f"{base}/instances", headers=hdrs, json={}, timeout=60)
        if r2.status_code == 401:
            raise PermissionError("Invalid or expired session ID.")
        r2.raise_for_status()
        instance_url = r2.json().get("url", "")
        if not instance_url:
            raise ValueError("Salesforce returned no instance URL.")

        poll_url = f"{SF_INSTANCE}{instance_url}"
        time.sleep(3)                 # give SF a moment to start the job
        deadline = time.time() + timeout_s
        attempt  = 0
        while time.time() < deadline:
            attempt += 1
            pr = requests.get(poll_url, headers=hdrs, timeout=60)
            pr.raise_for_status()
            pdata  = pr.json()
            status = pdata.get("status", "Running")
            if status_fn:
                status_fn(f"Async poll {attempt} — {status}")
            if status == "Success":
                return parse_sf_report(pdata)
            if status == "Failed":
                raise RuntimeError(
                    f"Report {report_id} failed: {pdata.get('errorCode', 'Unknown')}"
                )
            time.sleep(4)

        # Async timed out — surface partial sync data with a warning
        if sync_rows:
            if status_fn:
                status_fn(
                    f"⚠ Async timed out — using partial sync data ({len(sync_rows)} rows). "
                    f"Account-volume counts may be understated for large segments."
                )
            return sync_rows
        raise TimeoutError(
            f"Report {report_id} timed out after {timeout_s}s. "
            f"Try the CSV upload option below, or re-run after a few minutes."
        )

    except (PermissionError, RuntimeError):
        raise
    except TimeoutError:
        raise
    except Exception as async_err:
        if sync_rows:
            return sync_rows          # partial data is better than nothing
        raise RuntimeError(
            f"Could not fetch report {report_id}: {async_err}. "
            f"Try the CSV upload option below."
        )

# ─────────────────────────────────────────────────────────────────────────────
# ALGORITHM
# ─────────────────────────────────────────────────────────────────────────────
def parse_arr(val) -> float:
    if val is None or val == "":
        return 0.0
    try:
        return float(str(val).replace(",", "").replace("$", "").strip())
    except ValueError:
        return 0.0

def build_pools(vol_rows: list, new_accounts: list) -> dict:
    cse_count  = defaultdict(int)
    cse_arr    = defaultdict(float)
    cse_sub75k = defaultdict(int)

    for row in vol_rows:
        owner = norm(row.get("Account Owner") or "")
        if not owner or owner == "SMN Digital Commerce":
            continue
        arr = parse_arr(row.get("Contractual ARR"))
        cse_count[owner]  += 1
        cse_arr[owner]    += arr
        if arr < 7500:
            cse_sub75k[owner] += 1

    pools = {}
    for segment in ("Key", "Strategic", "Premier"):
        seg_cses = [c for c in CSE_BY_NORM.values() if c["segment"] == segment]
        active   = [c["norm"] for c in seg_cses if cse_count[c["norm"]] >= MIN_ACCOUNTS]
        if not active:
            pools[segment] = {}
            continue

        n_new        = sum(1 for a in new_accounts if a.get("segment") == segment)
        total_curr   = sum(cse_count[n] for n in active)
        target_avg   = (total_curr + n_new) / len(active)
        team_avg_arr = sum(cse_arr[n] for n in active) / len(active)

        pool = {}
        for n in active:
            ct, ar, s75 = cse_count[n], cse_arr[n], cse_sub75k[n]
            pool[n] = {
                "id":              CSE_BY_NORM[n]["id"],
                "raw_name":        CSE_BY_NORM[n]["name"],
                "current_count":   ct,
                "pre_run_count":   ct,   # frozen snapshot for summary
                "current_arr":     ar,
                "sub75k_count":    s75,
                "sub75k_ratio":    (s75 / ct) if ct else 0.0,
                "accounts_needed": target_avg - ct,
                "arr_gap":         team_avg_arr - ar,
                "new_count":       0,
            }
        pools[segment] = pool
    return pools

_CS_TEAM_MAP = {"key": "Key", "strategic": "Strategic", "premier": "Premier"}

def preprocess_accounts(rows: list):
    valid, bad = [], []
    for row in rows:
        cs_raw  = (row.get("CS Team") or "").strip()
        segment = next((v for k, v in _CS_TEAM_MAP.items() if k in cs_raw.lower()), None)
        if not segment:
            bad.append({**row, "_exception": "MISSING_CS_TEAM",
                        "_exception_detail": f"Unrecognized CS Team: '{cs_raw}'"})
            continue
        arr_raw = row.get("Contractual ARR (converted)") or row.get("Contractual ARR") or ""
        if arr_raw == "" or arr_raw is None:
            bad.append({**row, "_exception": "MISSING_ARR",
                        "_exception_detail": "Contractual ARR is blank"})
            continue
        arr = parse_arr(arr_raw)
        csm_raw = row.get("SecondAccountOwner") or row.get("Second Account Owner") or ""
        row = dict(row)
        row.update({"segment": segment, "ob_csm": norm_csm(csm_raw),
                    "arr": arr, "is_sub75k": arr < 7500})
        valid.append(row)
    return valid, bad

def _eligible(pool: dict) -> list:
    return [n for n, s in pool.items() if s["accounts_needed"] > 0]

def _update(pool: dict, n: str, arr: float, is_sub75k: bool):
    s = pool[n]
    s["current_count"]  += 1
    s["current_arr"]    += arr
    s["new_count"]      += 1
    s["accounts_needed"] -= 1
    if is_sub75k:
        s["sub75k_count"] += 1
    s["sub75k_ratio"] = s["sub75k_count"] / s["current_count"]

def _pick_key(account: dict, pool: dict, priority: str | None, p_done: bool) -> str | None:
    elig = _eligible(pool)
    if not elig:
        return None
    if priority and not p_done and priority in elig:
        return priority
    partnered = get_partnered_cses("Key", account["ob_csm"])
    if account["is_sub75k"]:
        return sorted(elig, key=lambda n: (-pool[n]["accounts_needed"], pool[n]["sub75k_ratio"], n))[0]
    pe = [n for n in elig if n in partnered]
    src = pe if pe else elig
    return sorted(src, key=lambda n: (-pool[n]["accounts_needed"], -pool[n]["sub75k_ratio"], n))[0]

def _pick_sp(account: dict, pool: dict, segment: str, priority: str | None, p_done: bool) -> str | None:
    elig = _eligible(pool)
    if not elig:
        return None
    if priority and not p_done and priority in elig:
        return priority
    partnered = get_partnered_cses(segment, account["ob_csm"])
    pe  = [n for n in elig if n in partnered]
    src = pe if pe else elig
    return sorted(src, key=lambda n: (-pool[n]["accounts_needed"], -pool[n]["arr_gap"], n))[0]

def run_assignment(new_rows: list, vol_rows: list, overrides: dict) -> dict:
    valid, bad = preprocess_accounts(new_rows)
    pools = build_pools(vol_rows, valid)

    results, exceptions = [], list(bad)
    run_summary = {}

    for segment in ("Key", "Strategic", "Premier"):
        pool      = pools.get(segment, {})
        seg_accts = [a for a in valid if a["segment"] == segment]
        ov_map    = overrides.get(segment, {})
        ov_given  = {n: 0 for n in ov_map}
        assigned_seg = excepted_seg = 0

        for account in seg_accts:
            priority = next(
                (n for n, min_ct in ov_map.items() if ov_given.get(n, 0) < min_ct), None
            )
            p_done = priority is None

            if segment == "Key":
                winner = _pick_key(account, pool, priority, p_done)
            else:
                winner = _pick_sp(account, pool, segment, priority, p_done)

            if winner:
                _update(pool, winner, account["arr"], account["is_sub75k"])
                if winner in ov_given:
                    ov_given[winner] += 1
                row = dict(account)
                row["New Account Owner"]    = pool[winner]["raw_name"]
                row["New Account Owner ID"] = pool[winner]["id"]
                row["_winner_norm"]         = winner
                row["_override"]            = (winner in ov_map and ov_given.get(winner, 0) == 1)
                results.append(row)
                assigned_seg += 1
            else:
                exceptions.append({**account, "_exception": "NO_ELIGIBLE_CSE",
                                   "_exception_detail": "All CSEs at or above target."})
                excepted_seg += 1

        cse_summary = []
        for n, s in sorted(pool.items(), key=lambda x: x[0]):
            cse_summary.append({
                "CSE":             s["raw_name"],
                "Before":          s["pre_run_count"],
                "Assigned":        s["new_count"],
                "After":           s["current_count"],
                "Total ARR":       round(s["current_arr"]),
            })
        run_summary[segment] = {"assigned": assigned_seg, "excepted": excepted_seg,
                                "cse_rows": cse_summary}

    exc_table = [{
        "Account ID":   e.get("18 Digit Account ID") or e.get("Account ID", ""),
        "Account Name": e.get("Account Name", ""),
        "CS Team":      e.get("CS Team", ""),
        "OB CSM":       e.get("SecondAccountOwner") or e.get("ob_csm", ""),
        "ARR":          e.get("Contractual ARR (converted)") or e.get("Contractual ARR", ""),
        "Code":         e.get("_exception", ""),
        "Detail":       e.get("_exception_detail", ""),
    } for e in exceptions]

    return {
        "results":    results,
        "exceptions": exc_table,
        "summary":    run_summary,
        "totals":     {"total": len(new_rows), "assigned": len(results), "excepted": len(exc_table)},
    }

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
_BLUE   = "0070F2"
_LBLUE  = "E1F4FF"
_WHITE  = "FFFFFF"
_RED_BG = "FDECEA"
_GREEN  = "E8F5E9"
_YELLOW = "FFF9C4"

def _hdr(cell, color=_BLUE):
    cell.font      = Font(bold=True, color=_WHITE, name="Calibri", size=10)
    cell.fill      = PatternFill("solid", fgColor=color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def build_excel(payload: dict) -> bytes:
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Assignments"

    results  = payload.get("results", [])
    _internal = {"segment","ob_csm","arr","is_sub75k","_winner_norm","_exception",
                 "_exception_detail","_override"}
    new_cols  = ["New Account Owner", "New Account Owner ID"]

    if results:
        base_cols = [k for k in results[0] if k not in _internal and k not in new_cols]
        all_cols  = base_cols + new_cols
        for ci, col in enumerate(all_cols, 1):
            cell = ws.cell(row=1, column=ci, value=col)
            _hdr(cell)
            ws.column_dimensions[get_column_letter(ci)].width = max(len(col) + 4, 16)
        for ri, row in enumerate(results, 2):
            override_row = row.get("_override", False)
            for ci, col in enumerate(all_cols, 1):
                c = ws.cell(row=ri, column=ci, value=row.get(col, ""))
                if col in new_cols:
                    c.fill = PatternFill("solid", fgColor=_YELLOW if override_row else _LBLUE)
        ws.freeze_panes = "A2"

    # ── Exceptions sheet ─────────────────────────────────────────────────────
    ws2  = wb.create_sheet("Exceptions")
    excs = payload.get("exceptions", [])
    if excs:
        exc_cols = list(excs[0].keys())
        for ci, col in enumerate(exc_cols, 1):
            cell = ws2.cell(row=1, column=ci, value=col)
            _hdr(cell)
            ws2.column_dimensions[get_column_letter(ci)].width = max(len(col) + 4, 16)
        for ri, row in enumerate(excs, 2):
            for ci, col in enumerate(exc_cols, 1):
                c = ws2.cell(row=ri, column=ci, value=row.get(col, ""))
                c.fill = PatternFill("solid", fgColor=_RED_BG)
    else:
        ws2.cell(row=1, column=1, value="No exceptions — all accounts assigned.")

    # ── Per-CSE Summary sheet ─────────────────────────────────────────────────
    ws3 = wb.create_sheet("CSE Summary")
    summary_cols = ["Segment", "CSE", "Before", "Assigned", "After", "Total ARR"]
    for ci, col in enumerate(summary_cols, 1):
        cell = ws3.cell(row=1, column=ci, value=col)
        _hdr(cell)
        ws3.column_dimensions[get_column_letter(ci)].width = max(len(col) + 4, 14)
    ri = 2
    for seg, sdata in payload.get("summary", {}).items():
        for row in sdata.get("cse_rows", []):
            ws3.cell(row=ri, column=1, value=seg)
            ws3.cell(row=ri, column=2, value=row["CSE"])
            ws3.cell(row=ri, column=3, value=row["Before"])
            ws3.cell(row=ri, column=4, value=row["Assigned"])
            ws3.cell(row=ri, column=5, value=row["After"])
            ws3.cell(row=ri, column=6, value=row["Total ARR"])
            if row["Assigned"] > 0:
                for ci in range(1, 7):
                    ws3.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=_GREEN)
            ri += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="tool-header">
  <div>
    <h1>CSE Account Assignment Tool</h1>
    <p>Connects to Salesforce &nbsp;·&nbsp; Applies RSE → CSE balancing algorithm &nbsp;·&nbsp; Exports to Excel</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar — settings ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Settings")
    st.caption("Monthly overrides give the named CSE their first account before normal rules apply.")
    st.markdown("**Strategic**")
    tom_override = st.checkbox("Tom Wahl — min. 1 account", value=True)
    st.markdown("**Premier**")
    alex_override = st.checkbox("Alex Capeloto — min. 1 account", value=True)

    st.divider()
    st.markdown("### Salesforce login")
    st.markdown("""
Sign in with your standard Salesforce credentials plus a **Security Token**.

**To get your security token:**
1. In Salesforce, click your profile icon → **Settings**
2. Go to **My Personal Information** → **Reset My Security Token**
3. Salesforce emails it to your address on file

Your credentials are sent directly to Salesforce and never stored.
""")
    st.divider()
    st.markdown("### CSV fallback")
    st.markdown("""
If login is unavailable, switch to **Upload CSV Files** mode.

For each report, open it in Salesforce → **Export** → **Details Only** → **Formatted Report: CSV**.
""")

# ── Overrides dict from sidebar ───────────────────────────────────────────────
active_overrides = {
    "Strategic": {"Thomas Wahl": 1} if tom_override else {},
    "Premier":   {"Alex Capeloto": 1} if alex_override else {},
}

# ── Input mode toggle ────────────────────────────────────────────────────────
input_mode = st.radio(
    "Data source",
    ["Salesforce Login", "Upload CSV Files"],
    horizontal=True,
    label_visibility="collapsed",
)

new_rows = None
vol_rows = None

if input_mode == "Salesforce Login":
    st.markdown(
        '<div class="info-box">'
        'Sign in with your Salesforce email, password, and security token. '
        'This creates an API session (not tied to your browser IP) so fetches work reliably from any server. '
        'See the sidebar for how to get your security token.'
        '</div>',
        unsafe_allow_html=True,
    )

    col_email, col_pass, col_token = st.columns([2, 2, 1.5])
    with col_email:
        sf_email = st.text_input("Salesforce Email", placeholder="you@company.com",
                                  label_visibility="visible")
    with col_pass:
        sf_password = st.text_input("Password", type="password",
                                     label_visibility="visible")
    with col_token:
        sf_token = st.text_input("Security Token", type="password",
                                  label_visibility="visible",
                                  help="Find this in Salesforce → Settings → My Personal Information → Reset My Security Token")

    col_test, col_run, col_spacer = st.columns([1.4, 1.4, 5])
    with col_test:
        test_clicked = st.button("Test Connection", use_container_width=True)
    with col_run:
        run_clicked = st.button("Fetch & Assign", type="primary", use_container_width=True)

    def _get_credentials():
        """Validate and return (email, password, token). Shows error and stops if blank."""
        if not sf_email.strip() or not sf_password.strip() or not sf_token.strip():
            st.error("Enter your Salesforce email, password, and security token before continuing.")
            st.stop()
        return sf_email.strip(), sf_password.strip(), sf_token.strip()

    if test_clicked:
        email, pwd, tok = _get_credentials()
        with st.spinner("Logging in to Salesforce…"):
            try:
                sid, _ = sf_login(email, pwd, tok)
            except (PermissionError, TimeoutError, RuntimeError) as e:
                st.error(f"**Login failed:** {e}")
                st.stop()

        with st.spinner("Testing report access…"):
            result = test_connection(sid)

        for step in result["steps"]:
            icon = "**:green[PASS]**" if step["ok"] else "**:red[FAIL]**"
            st.markdown(f"{icon} &nbsp; **{step['label']}** — {step['detail']}")

        if result["ok"]:
            st.success("All checks passed. Click **Fetch & Assign** to run the assignment.")
        else:
            st.warning("One or more checks failed — review the details above.")

    if run_clicked:
        email, pwd, tok = _get_credentials()

        try:
            with st.status("Connecting to Salesforce…", expanded=True) as sf_status:
                _sflog = lambda msg: sf_status.update(label=msg)

                sf_status.update(label="Logging in…")
                sid, _ = sf_login(email, pwd, tok)

                sf_status.update(label="Fetching New Accounts report…")
                new_rows = fetch_report(sid, RPT_NEW_ACCOUNTS, status_fn=_sflog)

                sf_status.update(
                    label=f"New Accounts: {len(new_rows)} rows.  Fetching Account Volumes…"
                )
                vol_rows = fetch_report(sid, RPT_VOLUMES, status_fn=_sflog)

                sf_status.update(
                    label=f"Volumes: {len(vol_rows)} rows.  Running algorithm…",
                    state="running",
                )
                payload = run_assignment(new_rows, vol_rows, active_overrides)
                sf_status.update(label="Complete.", state="complete")

            st.session_state["payload"]    = payload
            st.session_state["new_rows_n"] = len(new_rows)
            st.session_state["vol_rows_n"] = len(vol_rows)

        except PermissionError as e:
            st.error(
                f"**Authentication failed:** {e}  \n\n"
                "Check your email, password, and security token and try again. "
                "If you recently reset your password, request a new security token."
            )
            st.stop()
        except TimeoutError as e:
            st.error(
                f"**Timeout:** {e}  \n\n"
                "Try again in a minute, or switch to **Upload CSV Files** mode."
            )
            st.stop()
        except Exception as e:
            st.error(f"**Unexpected error:** {e}")
            st.stop()

else:
    # ── CSV upload fallback ───────────────────────────────────────────────────
    st.markdown(
        '<div class="info-box">'
        'In Salesforce, open each report → <strong>Export</strong> → <strong>Details Only</strong> '
        '→ <strong>Formatted Report: CSV</strong>. Upload both files below.'
        '</div>',
        unsafe_allow_html=True,
    )
    cu1, cu2 = st.columns(2)
    with cu1:
        new_file = st.file_uploader(
            "New Accounts Transitioning (report 00O0e000005i4gg)",
            type=["csv", "xlsx"],
            key="upload_new",
        )
    with cu2:
        vol_file = st.file_uploader(
            "Current Account Volumes (report 00O7V000006IT4i)",
            type=["csv", "xlsx"],
            key="upload_vol",
        )

    run_clicked = st.button(
        "Run Assignment", type="primary",
        disabled=(new_file is None or vol_file is None),
    )

    if run_clicked:
        try:
            def _read_upload(f):
                if f.name.endswith(".xlsx"):
                    df = pd.read_excel(f, dtype=str)
                else:
                    df = pd.read_csv(f, dtype=str)
                df = df.fillna("")
                return df.to_dict(orient="records")

            new_rows = _read_upload(new_file)
            vol_rows = _read_upload(vol_file)

            with st.spinner("Running assignment algorithm…"):
                payload = run_assignment(new_rows, vol_rows, active_overrides)

            st.session_state["payload"]    = payload
            st.session_state["new_rows_n"] = len(new_rows)
            st.session_state["vol_rows_n"] = len(vol_rows)

        except Exception as e:
            st.error(f"**Error processing uploaded files:** {e}")
            st.stop()


# ── Results display ───────────────────────────────────────────────────────────
if "payload" in st.session_state:
    payload = st.session_state["payload"]
    totals  = payload["totals"]
    summary = payload["summary"]
    results = payload["results"]
    exceptions = payload["exceptions"]

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Run Summary</div>', unsafe_allow_html=True)
    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    mc1.markdown(f'<div class="metric-card"><div class="label">Accounts Fetched</div>'
                 f'<div class="value">{totals["total"]}</div></div>', unsafe_allow_html=True)
    mc2.markdown(f'<div class="metric-card"><div class="label">Assigned</div>'
                 f'<div class="value green">{totals["assigned"]}</div></div>', unsafe_allow_html=True)
    mc3.markdown(f'<div class="metric-card"><div class="label">Exceptions</div>'
                 f'<div class="value orange">{totals["excepted"]}</div></div>', unsafe_allow_html=True)
    mc4.markdown(f'<div class="metric-card"><div class="label">Vol. Rows Loaded</div>'
                 f'<div class="value">{st.session_state.get("vol_rows_n","—")}</div></div>',
                 unsafe_allow_html=True)
    mc5.markdown(f'<div class="metric-card"><div class="label">Active Segments</div>'
                 f'<div class="value">{sum(1 for s in summary.values() if s["assigned"]>0)}</div></div>',
                 unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Download button ───────────────────────────────────────────────────────
    excel_bytes = build_excel(payload)
    st.download_button(
        label="⬇  Download Excel  (Assignments + Exceptions + CSE Summary)",
        data=excel_bytes,
        file_name="cse_assignments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_assign, tab_cse, tab_exc = st.tabs(
        [f"Assignments ({totals['assigned']})",
         "CSE Breakdown",
         f"Exceptions ({totals['excepted']})"]
    )

    with tab_assign:
        if results:
            _internal = {"segment","ob_csm","arr","is_sub75k","_winner_norm",
                         "_exception","_exception_detail","_override"}
            disp_cols = [k for k in results[0] if k not in _internal]
            df = pd.DataFrame(results)[disp_cols]

            # Highlight new-owner columns
            def _highlight(row):
                styles = [""] * len(row)
                for i, col in enumerate(row.index):
                    if col in ("New Account Owner", "New Account Owner ID"):
                        styles[i] = "background-color: #E1F4FF; font-weight: 600;"
                return styles

            st.dataframe(
                df.style.apply(_highlight, axis=1),
                use_container_width=True,
                hide_index=True,
                height=min(50 + 35 * len(df), 600),
            )
        else:
            st.info("No accounts were assigned in this run.")

    with tab_cse:
        for segment in ("Key", "Strategic", "Premier"):
            sdata = summary.get(segment, {})
            if not sdata.get("cse_rows"):
                continue
            with st.expander(
                f"**{segment}** — {sdata['assigned']} assigned this run",
                expanded=sdata["assigned"] > 0,
            ):
                cdf = pd.DataFrame(sdata["cse_rows"])

                def _highlight_assigned(row):
                    if row.get("Assigned", 0) > 0:
                        return ["background-color: #E8F5E9; font-weight: 600;"] * len(row)
                    return [""] * len(row)

                st.dataframe(
                    cdf.style.apply(_highlight_assigned, axis=1),
                    use_container_width=True,
                    hide_index=True,
                    height=min(50 + 35 * len(cdf), 500),
                )

    with tab_exc:
        if exceptions:
            edf = pd.DataFrame(exceptions)
            st.dataframe(edf, use_container_width=True, hide_index=True)
        else:
            st.success("No exceptions — every account was successfully assigned.")
