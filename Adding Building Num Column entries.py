

import sys
import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
SITE_URL       = "https://utoronto.sharepoint.com/sites/fs-projects"
USERNAME       = "nicolas.kondratenko@utoronto.ca"
PASSWORD       = "Kondrat2004!"         # ← ideally use an app-password or env-var
LIB_DISPLAY    = "Clean_Up_Test2"       # exact display title of your library
FIELD_INTERNAL = "Building"                   # internal name of the single-line text column “test”
TARGET_DEFAULT = "UNKNOWN"                # fallback if we can’t parse or find a mapping
TOP_N          = 5000                     # fetch up to 5000 items in one go

PREFIX_TO_BUILDING = {
    "P001": "001 - University College",
    "P002": "002 - Hart House",
    "P003": "003 - Gerstein Science Information Centre in the Sigmund Samuel Library Building",
    "P004": "004 - McMurrich Building",
    "P005": "005 - Medical Sciences Building",
    "P006": "006 - John P. Robarts Library Building",
    "P006A": "006A - Claude T. Bissell Building",
    "P006B": "006B - Thomas Fisher Rare Books Library",
    "P007": "007 - Lassonde Mining Building",
    "P008": "008 - Wallberg Building",
    "P008A": "008A - D.L. Pratt Building",
    "P009": "009 - Sandford Fleming Building",
    "P010": "010 - Simcoe Hall",
    "P010A": "010A - Convocation Hall",
    "P011": "011 - C. David Naylor Building",
    "P012": "012 - Munk School of Global Affairs at Trinity",
    "P014": "014 - Bloor Street West",
    "P016": "016 - Banting Institute",
    "P017": "017 - Queen's Park",
    "P018": "018 - Steam Plant",
    "P019": "019 - J. Robert S. Prichard Alumni House",
    "P020": "020 - Rosebrugh Building",
    "P021": "021 - Engineering Annex",
    "P022": "022 - Mechanical Engineering Building",
    "P023": "023 - University College Union",
    "P024": "024 - Haultain Building",
    "P025": "025 - FitzGerald Building",
    "P026": "026 - Cumberland House",
    "P027": "027 - St. George Street",
    "P028": "028 - Student Commons",
    "P030A": "030A - Varsity Centre",
    "P032": "032 - Wetmore Hall",
    "P033": "033 - Sidney Smith Hall",
    "P036": "036 - Astronomy & Astrophysics Building",
    "P038": "038 - Woodsworth College",
    "P040": "040 - Faculty of Law",
    "P041": "041 - Varsity Pavilion",
    "P042": "042 - Goldring Centre for High Performance Sport",
    "P043": "043 - School of Graduate Studies",
    "P047": "047 - Canadiana Gallery",
    "P050": "050 - Falconer Hall",
    "P051": "051 - Edward Johnson Building",
    "P053": "053 - Dr. Eric Jackman Institute of Child Study",
    "P054": "054 - Daniels Building",
    "P056": "056 - Graduate Students' Union",
    "P057": "057 - Bancroft Building",
    "P061": "061 - Borden Building South",
    "P061A": "061A - Borden Building North",
    "P062": "062 - Earth Sciences Centre",
    "P065": "065 - Dentistry Building",
    "P066": "066 - Spadina Avenue",
    "P067": "067 - Huron Street",
    "P068": "068 - Clara Benson Building",
    "P068A": "068A - Warren Stevens Building",
    "P070": "070 - Galbraith Building",
    "P072": "072 - Ramsay Wright Laboratories",
    "P073": "073 - Lash Miller Chemical Laboratories",
    "P075": "075 - Faculty Club",
    "P077": "077 - Sussex Court",
    "P078": "078 - McLennan Physical Laboratories",
    "P079": "079 - Anthropology Building",
    "P080": "080 - Bahen Centre for Information Technology",
    "P082": "082 - Gage Building",
    "P083": "083 - McCaul Street",
    "P087": "087 - Myhal Centre for Engineering Innovation & Entrepreneurship",
    "P088": "088 - St. George Street",
    "P089": "089 - Munk School of Global Affairs at the Observatory",
    "P090": "090 - College Street",
    "P091": "091 - Luella Massey Studio Theatre",
    "P092": "092 - Communications House",
    "P095": "095 - ",
    "P097A": "097A - Queen's Park Crescent. E.",
    "P098": "098 - Wellesley Street West",
    "P103": "103 - School of Continuing Studies",
    "P104": "104 - Max Gluskin House",
    "P105": "105 - Fields Inst for Research in Math Science",
    "P106": "106 - St. George Street",
    "P110": "110 - St. George Street",
    "P111": "111 - Factor",
    "P120": "120 - Louis B. Stewart Observatory",
    "P122": "122 - Northwest Chiller Plant",
    "P123": "123 - Ontario Institute for Studies in Education",
    "P125": "125 - Spadina Avenue",
    "P127": "127 - Enrolment Services",
    "P128": "128 - Jackman Humanities Building",
    "P129": "129 - Early Learning Centre",
    "P130": "130 - Woodsworth College Residence",
    "P131": "131 - New College III",
    "P132": "132 - Innis College",
    "P134": "134 - Rotman School of Management",
    "P135": "135 - St. George Parking Garage",
    "P138": "138 - Huron Street",
    "P143": "143 - Koffler Student Services Centre",
    "P145": "145 - Koffler House",
    "P146": "146 - Sussex Avenue",
    "P151": "151 - Fasken Martineau Building",
    "P152": "152 - Rehabilitation Sciences Building",
    "P154": "154 - Health Sciences Building",
    "P155": "155 - Exam Centre",
    "P156": "156 - Old Admin Bldg",
    "P160": "160 - Terrence Donnelly Ctr for Cellular & Biomolecular Res",
    "P161": "161 - Leslie L. Dan Pharmacy Building",
    "P165": "165 -  ",
    "P171": "171 - Spadina Ave",
    "P172": "172 - Macdonald",
    "P174": "174 - Beverley Street",
    "P179": "179 - College Street",
    "P189": "189 - Spadina Avenue",
    "P192": "192 - Stewart Building",
    "P193": "193 - Edward Street",
    "P194": "194 - TWH ",
    "P195": "195 - MARS 2",
    "P196": "196 - University Avenue",
    "P197": "197 - Bay Street",
    "P197A": "197A -",
    "P197B": "197B -",
    "P197C": "197C -",
    "P197E": "197E -",
    "P200": "200 B/H/R/S/201 -",
    "P200B": "200B - Bladen Wing (B",
    "P200H": "200H - Humanities Wing (H",
    "P200JKL": "200JKL - Portable 101/102/103",
    "P200M": "200M - Social Sciences Building",
    "P200NP": "200NP - Portable 104/105",
    "P200R": "200R - Highland Hall",
    "P200S": "200S - Science Wing (S",
    "P201": "201 - Academic Resource Centre",
    "P203": "203 - UTSC Student Centre",
    "P204": "204 - Arts & Admin Building",
    "P205": "205 - Science Research Building",
    "P206": "206 - Instructional Centre (UTSC)",
    "P207": "207 - Environmental Science & Chemistry",
    "P230": "230 - Student Residence Centre",
    "P231": "231 - N'sheemaehn Child Care Centre",
    "P261": "261 -",
    "P263": "263 -",
    "P301": "301 -",
    "P302": "302 -",
    "P312": "312 -",
    "P313": "313 - William G. Davis Building",
    "P314": "314 - Kaneff Centre for Management & Social Sciences",
    "P314A": "314A - Innovation Complex",
    "P316": "316 - Erindale Studio Theatre",
    "P317": "317 -",
    "P322": "322 - Geomorphology",
    "P323R": "323R -",
    "P328": "328 - UTM Student Centre",
    "P329": "329 - Communication Culture & Technology",
    "P330": "330 - UTM Alumni House",
    "P331": "331 - Hazel McCallion Academic Learning Centre",
    "P332": "332 - Recreation Athletics & Wellness",
    "P333": "333 - Terrence Donnelly Health Sciences Complex",
    "P334": "334 - Instructional Centre (UTM)",
    "P335": "335 - Academic Annex",
    "P336": "336 - UTM Grounds Building",
    "P338": "338 - Research Greenhouse",
    "P340": "340 - Deerfield Hall",
    "P341": "341 - Maanjiwe nendamowinan",
    "P405": "405 - Elmsley Hall",
    "P407": "407 - Muzzo Family Alumni Hall",
    "P410": "410 -",
    "P411": "411 - Brennan Hall",
    "P415": "415 - Odette (Louis) Hall",
    "P416": "416 - Windle House",
    "P426": "426 - Carr Hall",
    "P428": "428 - Founders House",
    "P429": "429 - J.M. Kelly Library",
    "P501": "501 - Victoria College",
    "P502": "502 - Emmanuel College",
    "P503": "503 - Birge",
    "P504": "504 - Burwash Hall",
    "P507": "507 - Goldring Student Centre",
    "P509": "509 - Isabel Bader Theatre",
    "P513": "513 - Stephenson House",
    "P514": "514 - E.J. Pratt Library",
    "P515": "515 - Northrop Frye Hall",
    "P516": "516 - Charles Street West",
    "P528": "528 - Lillian Massey Building",
    "P602": "602 - Gerald Larkin Building",
    "P603": "603 - George Ignatieff Theatre",
    "P434": "434 - Toronto School of Theology",
    "P478": "478 - Regis College",
    "P049": "049 - Aerospace Building",
    "P149": "149 - UTL at Downsview",
    "P999": "999 - Multi-Building"
}

def main():

    #ATTEMPT TO ACESS THE SHAREPOINT

    print("Authenticating to Sharepoint")

    try:
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
    except Exception as e:
        print(f"‼️ Authentication failed: {e}")
        sys.exit(1)

    print(f"▶ Loading library “{LIB_DISPLAY}” …")
    sp_list = ctx.web.lists.get_by_title(LIB_DISPLAY)
    ctx.load(sp_list, ["Title", "RootFolder"])
    ctx.execute_query()
    print(f"   • Confirmed library title = '{sp_list.properties['Title']}'")

    root_folder = sp_list.root_folder
    ctx.load(root_folder, ["ServerRelativeUrl", "Folders"])
    ctx.load(root_folder.folders)
    ctx.execute_query()
    first_level = root_folder.folders

    if not first_level:
        print("‼️ No first-level folders found under the library root.")
        return
    
    print(f"▶ Found {len(first_level)} first-level folder(s):")
    for f in first_level:
        print(f"   • {f.properties.get('Name')}")
    
    update = 0
    print(f"\n▶ Updating “{FIELD_INTERNAL}” for each first-level folder…")

    for folder in first_level:
        folder_name = folder.properties.get("Name")
        prefix = folder_name.split("-", 1)[0].strip()
        mapped = PREFIX_TO_BUILDING.get(prefix, TARGET_DEFAULT)

        item = folder.list_item_all_fields
        ctx.load(item, [FIELD_INTERNAL])
        ctx.execute_query()

        old_val = item.properties.get(FIELD_INTERNAL)
        print(f"\n   → Folder = {folder_name!r}")
        print(f"     Prefix = {prefix!r}")
        print(f"     (Before) {FIELD_INTERNAL!r} = {old_val!r}")
        print(f"     (Mapped) {mapped!r} → writing into {FIELD_INTERNAL!r}")

        try:
            item.set_property(FIELD_INTERNAL,mapped)
            item.update()
            ctx.execute_query()
            print(f"     ✔ (After)  {FIELD_INTERNAL!r} = {mapped!r}")
            updated += 1
        except Exception as e:
            print(f"     ✘ Skipped {folder_name!r} due to error: {e}")

    print(f"\n✅ Done. {updated}/{len(first_level)} folder(s) updated.\n")



if __name__ == "__main__":

    try:
        main()
    except Exception as err:
        print(f" ‼️ Fatal error: {err}")
        sys.exit(1)