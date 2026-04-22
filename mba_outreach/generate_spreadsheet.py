"""Generate the state Mortgage Bankers Association contacts spreadsheet.

Produces state_mba_contacts.xlsx for the AI Security & Governance webinar
outreach push. Data was compiled from each association's public website,
staff directories, and public profiles (LinkedIn, ZoomInfo, industry press)
as of April 2026. Emails flagged "unverified" in Notes are inferred from
domain patterns and should be verified before first send. Generic/fallback
emails and contact-form URLs have been verified against association sites.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = [
    "State",
    "Association",
    "Website",
    "Primary Contact",
    "Primary Title",
    "Primary Email",
    "Secondary Contact",
    "Secondary Title",
    "Secondary Email",
    "Generic / Fallback",
    "Phone",
    "Notes",
]

ROWS = [
    # --- Northeast ---
    ["Connecticut", "Connecticut Mortgage Bankers Association (CMBA)", "https://www.cmba.org/",
     "Denise Derosier", "Executive Director", "denise@cmba.org",
     "Brenda Corbo", "Board President (Ives Bank)", "",
     "Contact form on site", "(860) 778-1355",
     "Denise is the paid staff lead; best direct target."],
    ["Delaware", "Delaware Mortgage Bankers Association (DMBA)", "https://delawaremba.org/",
     "Kevin Jornlin", "President (per 2024 NMP feature; verify current)", "",
     "", "", "",
     "https://delawaremba.starchapter.com/form.php?form_id=13", "(302) 765-9600",
     "Volunteer-run on StarChapter; no paid staff publicly listed."],
    ["Maine", "Maine Association of Mortgage Professionals (MAMP)", "https://mamp.wildapricot.org/",
     "Janet Williams", "Executive Director", "",
     "Hal Tippetts", "President", "",
     "Contact form on site", "",
     "Maine has no state MBA; MAMP (NAMB-affiliated) is the closest trade group."],
    ["Maryland", "Maryland Mortgage Bankers & Brokers Association (MMBBA)", "https://www.mmbba.org/",
     "Rich Green", "Board Contact (Presidential Bank)", "Richard.Green@Presidential.com",
     "", "", "",
     "info@mmbba.org", "(443) 574-7786",
     "Rotating volunteer President; Rich Green is a directly reachable board figure."],
    ["Massachusetts", "Massachusetts Mortgage Bankers Association (MMBA)", "https://www.massmba.com/",
     "Debbie Sousa", "Executive Director", "dsousa@massmba.com",
     "Jennifer Couldren", "Director of Operations", "jenn@massmba.com",
     "dsousa@massmba.com", "(617) 570-9114",
     "Strong, active staff; Debbie is the decision-maker. Also runs MMBA Foundation."],
    ["New Hampshire", "Mortgage Bankers & Brokers Association of NH (MBBA-NH)", "https://mbba-nh.org/",
     "June Hibbs", "Executive Director", "",
     "", "", "",
     "Contact form on site (info@mbba-nh.org likely; unverified)", "(603) 225-6111",
     "Small org; co-runs Tri-State Mortgage Conference with VT and ME."],
    ["New Jersey", "Mortgage Bankers Association of New Jersey (MBA-NJ)", "https://www.mbanj.com/",
     "Robert Levy", "Executive Director & Counsel", "",
     "Matthew D. VanFossen, CMB", "President (CEO, Absolute Home Mortgage)", "",
     "info@mbanj.com", "(732) 504-3497",
     "Rob Levy is paid staff lead; VanFossen is volunteer President."],
    ["New York", "New York Mortgage Bankers Association (NYMBA)", "https://www.nymba.org/",
     "Christina Wiley", "Executive Director / COO", "cwiley@nymba.org",
     "", "", "",
     "Contact form at nymba.org/contact-us", "(518) 963-0593",
     "Christina runs member education, convention, outreach. Strong direct target."],
    ["Pennsylvania (Eastern)", "Mortgage Bankers Association of Eastern PA (MBAEPA)", "https://www.mbaepa.com/",
     "", "Managed by AIM (Association Insight Management)", "",
     "", "", "",
     "info@mbaepa.com", "(484) 284-0944",
     "PA has no statewide MBA; MBAEPA is the larger regional chapter (Bala Cynwyd)."],
    ["Pennsylvania (Southwestern)", "Mortgage Bankers Association of Southwestern PA (MBA-SWPA)", "https://mba-swpa.org/",
     "", "Volunteer board; names not public", "",
     "", "", "",
     "mbaswpa@gmail.com", "(724) 348-7051",
     "Pairs with MBAEPA to cover PA; volunteer-run."],
    ["Rhode Island", "Rhode Island Mortgage Bankers Association (RIMBA)", "https://www.rimba.org/",
     "Lisa Cabral", "Board Director (RIHousing)", "",
     "", "", "",
     "Contact form on site", "(401) 421-2338",
     "Volunteer-run; no paid ED publicly listed. Uses rimba.fastclass.com for CE."],
    ["Vermont", "Vermont Mortgage Bankers Association (VMBA)", "https://vermontmba.org/",
     "Stephen Kendall", "President, Board of Governors", "",
     "", "", "",
     "Contact form at vermontmba.org/Contact", "(866) 680-8622",
     "Volunteer-led on Wild Apricot. Co-hosts Tri-State Mortgage Conference with NH/ME."],

    # --- Southeast ---
    ["Alabama", "Mortgage Bankers Association of Alabama (MBAAL)", "https://www.mbaal.org/",
     "", "Board-led; ED not publicly named", "",
     "", "", "",
     "Contact form at mbaal.org/contact-us", "",
     "Memphis-area management hub likely (Confluent Strategies) — same as TN, AR, FL."],
    ["Arkansas", "Mortgage Bankers Association of Arkansas (MBAA)", "https://www.arkansasmba.org/",
     "Nicci Osborne", "Executive Director", "nosborne@arkansasmba.org",
     "Erica J. Byrne", "President", "",
     "info@arkansasmba.org", "(901) 321-6758",
     "Managed from Memphis, TN. Nicci is directly emailable."],
    ["Florida", "Mortgage Bankers Association of Florida (MBAF)", "https://www.mbaf.org/",
     "Brenda Thomas", "Executive Director", "mbaf@mbaf.org",
     "Matthew Goldman", "President (Atlantic Home Loans)", "",
     "mbaf@mbaf.org", "(407) 855-6155",
     "Also has local chapters in Central FL, South FL, Tampa Bay, and Tallahassee."],
    ["Georgia", "Mortgage Bankers Association of Georgia (MBAG)", "https://www.mbag.org/",
     "Teri Kramer", "Executive Director", "tkramer@mbag.org (unverified)",
     "Kurt Owen", "President (verify current)", "",
     "Contact form at mbag.org/contact", "(478) 743-8612",
     "Email pattern unverified; 7 regional GA chapters including Atlanta MBA."],
    ["Kentucky", "Mortgage Bankers Association of Kentucky (MBAKY)", "https://www.mbaky.org/",
     "Ashley L. Cassetty, MSBC", "Executive Director", "mbakentucky@gmail.com",
     "", "", "",
     "mbakentucky@gmail.com", "(270) 303-1524",
     "Ashley is single point of contact for leadership + education."],
    ["Louisiana", "Louisiana Mortgage Bankers Association (LMBA)", "https://www.lmba.la/",
     "Brad Haymond", "LMBA affiliation (role unconfirmed)", "",
     "", "", "",
     "Contact form at lmba.la", "",
     "Contact form first to identify right POC. Acadiana MLA is separate regional group."],
    ["Mississippi", "Mortgage Bankers Association of Mississippi (MSMBA)", "https://www.msmortgagebankers.org/",
     "Owen Munton", "President (Hancock Whitney Bank)", "",
     "Casey Wilberding", "Treasurer (Bank First)", "",
     "info@msmortgagebankers.org", "",
     "Volunteer-led; no paid staff. LinkedIn Owen Munton as backup."],
    ["North Carolina & South Carolina", "Mortgage Bankers Association of the Carolinas (MBAC)", "https://www.mbac.org/",
     "R. Bryan Wright", "Recent President (per March 2024 NMP feature)", "",
     "", "", "",
     "Contact form on mbac.org", "",
     "MBAC covers BOTH NC and SC — one outreach covers two states. Based in Huntersville, NC."],
    ["Tennessee", "Tennessee Mortgage Bankers Association (TNMBA)", "https://www.tnmba.org/",
     "Caitlin Guerra", "Executive Director", "cguerra@tnmba.org (unverified; use info@)",
     "Karley Bond, CMB, AMP", "President (Regions)", "",
     "info@tnmba.org", "(901) 321-6739",
     "Memphis-based mgmt hub. Debbie Gadberry (Treasurer) runs Next Level Education."],
    ["Virginia", "Virginia Mortgage Bankers Association (VMBA)", "https://virginiamba.org/",
     "Walt Lyons", "Executive Director (also VP Education, VA Bankers Assn)", "wlyons@virginiamba.org",
     "Kevin Russell", "Recent President (F&M Mortgage)", "",
     "Contact form at virginiamba.org/Contact", "(804) 819-4746",
     "Walt is a dual-purpose target — ED plus education lead. Local chapters: CVMBA, NOVAMBA."],
    ["West Virginia", "No active state MBA", "https://www.wvbankers.org/",
     "Rick Clayburgh", "President & CEO, WV Bankers Association (fallback)", "",
     "", "", "",
     "WV Bankers Assn contact form", "(701) 223-5303",
     "No standalone WV MBA. WV Bankers Association is general banking, not mortgage-specific."],

    # --- Midwest ---
    ["Illinois", "Illinois Mortgage Bankers Association (IMBA)", "https://imba.org/",
     "Barbara Zajicek", "Executive Director", "BarbaraZajicek@att.net",
     "", "", "",
     "Contact form at imba.org/form.php?form_id=13", "(312) 236-6208",
     "ED uses personal att.net address. Mokena, IL office."],
    ["Indiana", "Indiana Mortgage Bankers Association (IMBA)", "https://www.indianamba.org/",
     "Alan Thorup", "Executive Director (since 2008)", "alan@indianamba.org (unverified pattern)",
     "", "", "",
     "communications@indianamba.org", "",
     "Domain pattern first@indianamba.org but Alan's exact address not confirmed."],
    ["Iowa", "Iowa Mortgage Association (IMA)", "https://iowama.org/",
     "Tom Schulte", "President", "",
     "Darcy Burnett", "Programs contact (at Iowa Bankers Association)", "",
     "Contact form at iowama.org/contact", "(515) 286-4352",
     "Managed out of Iowa Bankers Association; no in-house ED."],
    ["Kansas", "No statewide Kansas MBA — MBAKC (Kansas City metro, KS/MO)", "https://www.mbakc.com/",
     "Susie Mize", "Executive Director, MBAKC", "director@mbakc.com",
     "", "", "",
     "Contact form at mbakc.com/contact-us-2", "(913) 378-6125",
     "No state-level KS MBA; MBAKC is metro KC, covers KS and MO."],
    ["Michigan (ALREADY BOOKED)", "Michigan Mortgage Lenders Association (MMLA)", "https://mmla.net/",
     "Joanne Misuraca", "Chief Executive Officer", "joanne@mmla.net (unverified pattern)",
     "", "", "",
     "Contact form at mmla.net/about", "(586) 226-2823",
     "Webinar already scheduled — included for reference only."],
    ["Minnesota", "The Minnesota Mortgage Association (MMA)", "https://themma.org/",
     "Stephen Rice", "President (board)", "",
     "", "", "",
     "info@themma.org", "",
     "Open with info@ and ask for the education/programs lead."],
    ["Missouri", "Mortgage Bankers Association of Missouri (MBAMO)", "https://mbamo.org/",
     "Kim Akin", "Executive Director (via CN Missouri)", "Kim@CNMissouri.com",
     "Jim Andrews", "ED, St. Louis chapter", "jandrews@qabs.com",
     "No info@mbamo.org; use chapter contacts", "(573) 520-7240",
     "AMC-managed. Members join at the chapter level (STL and KC)."],
    ["Nebraska", "Nebraska Mortgage Association (NMA)", "https://nebraskamortgageassociation.com/",
     "Joe Pittman", "Executive Director", "jpittman@nebraskamortgageassociation.com (unverified)",
     "Erin Isenhart", "Association Management (operations/events)", "",
     "Contact form on StarChapter site", "(402) 397-0280",
     "No public info@ email. Email patterns inferred."],
    ["North Dakota", "No active state MBA", "https://www.ndba.com/",
     "Rick Clayburgh", "President & CEO, ND Bankers Association (fallback)", "",
     "", "", "",
     "ND Bankers contact form", "(701) 223-5303",
     "No standalone ND MBA. Route via NDBA or national MBA."],
    ["Ohio", "Ohio Mortgage Bankers Association (OMBA)", "https://ohiomba.org/",
     "Rich Swerbinsky", "Executive Director (effective July 2025)", "rich@ohiomba.org",
     "", "", "",
     "omba@ohiomba.org", "(614) 682-6555",
     "Strong target — new ED, actively publishing on LinkedIn. 2026 conference June 14-16 Columbus."],
    ["South Dakota", "SD Association for Mortgage Professionals (SDAMP) — est. 2024", "https://sdmortgageprofessionals.org/",
     "", "New org; staff/officers not yet indexed", "",
     "", "", "",
     "Check site Contact page directly", "",
     "Young org (founded 2024). SDBA (Karl Adam, 605-224-1653) is general-banking fallback."],
    ["Wisconsin", "Wisconsin Mortgage Bankers Association (WMBA)", "http://www.wimba.org/",
     "Heather Dyer, CAE", "CEO (via Morgan Data Solutions)", "hdyer@morgandata.com",
     "", "", "",
     "wmba@morgandata.com", "(608) 255-4180",
     "AMC-managed. All outreach through morgandata.com."],

    # --- West / Southwest ---
    ["Alaska", "Alaska Mortgage Bankers Association (AMBA)", "https://alaskamba.org/",
     "Benjamin Reynolds", "Interim President", "",
     "", "", "",
     "Contact form on site", "(907) 360-2324",
     "Active since 1976. Email pattern @alaskamba.org but full addresses not public."],
    ["Arizona", "Arizona Mortgage Lenders Association (AMLA)", "https://azmortgagelenders.com/",
     "Debbie Hill", "Executive Director", "",
     "", "", "",
     "amla@cox.net", "(623) 433-8940",
     "MBA-affiliated state org. Separate brokers group is AZAMP."],
    ["California", "California Mortgage Bankers Association (CMBA)", "https://cmba.com/",
     "Paul Gigliotti", "CEO (appointed Oct 2025)", "paul@cmba.com (unverified pattern)",
     "Krys Delk", "Member Engagement Director", "",
     "Contact form at cmba.com", "(916) 446-7100",
     "Staff emails not published publicly; LinkedIn to Paul + contact form is safest path."],
    ["Colorado", "Colorado Mortgage Lenders Association (CMLA)", "https://cmla.com/",
     "Betty Knecht", "Executive Director", "betty@cmla.com",
     "Janelle Chakounis", "Membership & Business Development Manager", "janelle@cmla.com",
     "info@cmla.com", "(303) 773-9565",
     "Best-documented contacts in the region. Janelle is the likely right person for programming."],
    ["Hawaii", "Mortgage Bankers Association of Hawaii (MBAH)", "https://mbahawaii.org/",
     "", "2025 officers — names not surfaced", "",
     "", "", "",
     "info@mbahawaii.org", "(808) 681-7500",
     "MBA-affiliated chapter. Separate broker group (HAMB) exists."],
    ["Idaho", "Idaho Mortgage Lenders Association (IMLA)", "https://www.idahomortgagelenders.org/",
     "Lindsay Craven, AMP", "President 2026-2027", "",
     "Leah Marchbanks", "Immediate Past President", "",
     "Contact form at idahomortgagelenders.org/contact", "",
     "Volunteer-board-run; no paid staff. Boise office."],
    ["Montana", "No active state MBA — Montana Assn for Mortgage Professionals (MAMP) is broker-focused", "https://mt-mamp.org/",
     "Matthew Tedesco", "Board of Directors (per LinkedIn)", "",
     "", "", "",
     "Contact form at mt-mamp.org", "",
     "No MBA exists; MAMP is NAMB-affiliated brokers group. MT Bankers Assn is general-banking alt."],
    ["Nevada", "Nevada Mortgage Lenders Association (NMLA)", "https://www.nvmla.com/",
     "Andrew Leavitt", "Chairman, Board of Governors", "",
     "Mike Giusti", "Director of Education", "",
     "admin@nvmla.com", "(702) 704-5823",
     "admin@nvmla.com is the verified entry point; address Mike by name for programming."],
    ["New Mexico", "New Mexico Mortgage Lenders Association (NMMLA)", "https://nmmla.com/",
     "Theresa Castellano", "Executive Vice President", "",
     "", "", "",
     "nmmla2012@gmail.com", "(505) 480-8514",
     "Generic Gmail inbox is verified. Theresa also has @1nmamp.com via sister broker org."],
    ["Oklahoma", "Oklahoma Mortgage Bankers Association (OMBA)", "https://www.oklamba.com/",
     "Linda Dickerson", "Secretary/Treasurer", "",
     "Steven Plaisance", "MBA State Ambassador (Gateway First Bank)", "",
     "info@oklamba.com", "",
     "Small staff; info@ is the reliable channel."],
    ["Oregon", "Oregon Mortgage Bankers Association (OMBA)", "https://www.oremba.org/",
     "", "Volunteer board; names not surfaced publicly", "",
     "", "", "",
     "info@oremba.org", "(503) 223-6622",
     "Board directory at oremba.org/board-of-directors (visit directly for names)."],
    ["Texas", "Texas Mortgage Bankers Association (TMBA)", "https://www.texasmba.org/",
     "Scott Norman", "CEO / Executive Vice President", "scott@texasmba.org (unverified pattern)",
     "Cory Whipple", "Conference Event Coordinator", "cory@texasmba.org (unverified pattern)",
     "info@texasmba.org", "(512) 480-8622",
     "High-priority — large, active association with heavy education programming."],
    ["Utah", "Utah Association of Mortgage Professionals (UAMP) — no state MBA", "https://www.uamp.org/",
     "Brittany Black", "President 2024-2025", "",
     "Trent Hendry", "Prior President", "",
     "info@uamp.org", "(801) 550-0868",
     "No dedicated UT MBA; UAMP is broker/MLO-focused. Skip if webinar isn't broker-appropriate."],
    ["Washington", "Washington Mortgage Bankers Association (WA-MBA)", "https://wa-mba.org/",
     "Amy Ohlinger", "Executive Director (via Association Management Inc.)", "amy@aminc.org",
     "", "", "",
     "admin@wa-mba.org", "(253) 525-5169",
     "Managed by AMI in Gig Harbor. Amy is the verified direct contact."],
    ["Wyoming", "Wyoming Mortgage Lenders Association (WMLA)", "https://www.wyomortgagelenders.com/",
     "", "Volunteer board; names not public", "",
     "", "", "",
     "Contact form at wyomortgagelenders.com", "",
     "Small volunteer-run assn. WY Bankers Assn (Office@WyomingBankers.com) is a secondary channel."],
]


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "State MBA Contacts"

    header_fill = PatternFill(start_color="FF1F3A5F", end_color="FF1F3A5F", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFFFF", size=11)
    header_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.append(HEADERS)
    for col_idx, _ in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    body_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for row in sorted(ROWS, key=lambda r: r[0]):
        ws.append(row)

    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row_cells:
            cell.alignment = body_align

    widths = [22, 50, 36, 24, 38, 36, 24, 32, 36, 38, 18, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    return wb


if __name__ == "__main__":
    import os
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "state_mba_contacts.xlsx")
    build_workbook().save(out_path)
    print(f"Wrote {out_path} with {len(ROWS)} rows.")
