"""
Add Broomfield County and Weld County rows to the Resources sheet.
Update the county dropdown to include all 4 counties.
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

DST = "/home/user/Appraisal-Template/Analysis_Workbook_OPTIMIZED.xlsm"

wb = load_workbook(DST, keep_vba=True)
ws = wb["Resources"]

# Priority formula template (same pattern as existing rows)
def prio(r):
    return (
        f'=IF(B{r}<>$D$3,"◌ Other",'
        f'IF(C{r}="REGIONAL / ALL MUNICIPALITIES","Regional",'
        f'IF(OR('
        f'ISNUMBER(SEARCH(LEFT(\'Property Info\'!$B$17,MAX(1,FIND(",",\'Property Info\'!$B$17&",")-1)),C{r})),'
        f'ISNUMBER(SEARCH(\'Property Info\'!$B$18,C{r}))'
        f')),"★ Active","–")))'
    )

URL_FONT = Font(color="0563C1", underline="single", size=10)
ROW_FILL = PatternFill("solid", fgColor="F5F7FA")

def append_row(county, municipality, category, name, url, description):
    r = ws.max_row + 1
    ws.cell(r, 1, prio(r))
    ws.cell(r, 2, county)
    ws.cell(r, 3, municipality)
    ws.cell(r, 4, category)
    ws.cell(r, 5, name)
    url_cell = ws.cell(r, 6, url)
    url_cell.hyperlink = url
    url_cell.font = URL_FONT
    ws.cell(r, 7, description)
    ws.row_dimensions[r].height = 15
    return r

# ── BROOMFIELD COUNTY ────────────────────────────────────────────────────────
BC = "Broomfield County"
BM = "Broomfield (City & County)"

broomfield_rows = [
    # Property Assessment
    (BC, BM, "Property Assessment", "Assessor Property Search Portal", "https://egov.broomfield.org/assessor/web/", "Official assessor database — search by owner, address, parcel, or legal description"),
    (BC, BM, "Property Assessment", "Parcel Map / GIS Viewer", "https://www.broomfield.org/2739/Parcel-Search", "Interactive GIS parcel map with zoning, land use, and assessed value data"),
    (BC, BM, "Property Assessment", "Valuation Process & Appeals", "https://www.broomfield.org/3634/Valuation-Process", "NOV deadlines, assessment process, and appeal procedures"),
    (BC, BM, "Property Assessment", "Appeal Your Property Valuation", "https://www.broomfield.org/4013/Appeal-Information", "Online portal to file valuation protests with documentation upload"),
    # Property Tax / Treasurer
    (BC, BM, "Property Tax / Treasurer", "Treasurer Online Portal", "https://egov.broomfield.org/treasurer/web/", "Tax bill search, tax notice printing, and online payment"),
    (BC, BM, "Property Tax / Treasurer", "Pay Property Taxes Online", "https://www.broomfield.org/2348/Pay-a-Bill-Online", "Online payment with EquaPay monthly installment option"),
    (BC, BM, "Property Tax / Treasurer", "Property Tax 101", "https://broomfield.org/4458/Property-Tax-101", "Mill levy, assessment rate, and tax calculation explained"),
    # Clerk & Recorder
    (BC, BM, "Clerk & Recorder / Official Records", "Recorder Online Portal", "https://egov.broomfield.org/recorder/web/", "Recorded documents portal — deeds, deeds of trust, and liens from Nov 2001"),
    (BC, BM, "Clerk & Recorder / Official Records", "Recording Division", "https://www.broomfield.org/470/Recording", "eRecording and in-person document recording services"),
    # GIS / Mapping
    (BC, BM, "GIS / Mapping", "Broomfield Open Data Portal", "https://opendata.broomfield.org/", "Parcel data, addresses, zoning, and GIS layers in CSV / shapefile / GeoJSON"),
    (BC, BM, "GIS / Mapping", "GIS Division & Data Downloads", "https://www.broomfield.org/1761/GIS-Division-Data", "Downloadable GIS datasets: parcels, zoning, floodplains, and more"),
    (BC, BM, "GIS / Mapping", "Maps for Download", "https://www.broomfield.org/225/Maps-for-Download", "Collection of static and interactive maps from Broomfield"),
    # Planning & Zoning
    (BC, BM, "Planning & Zoning", "Planning Division", "https://www.broomfield.org/290/Planning", "Land use planning, development review, and zoning administration"),
    (BC, BM, "Planning & Zoning", "Zoning Information", "https://www.broomfield.org/1979/Zoning", "Zoning districts, regulations, and permitted uses"),
    (BC, BM, "Planning & Zoning", "Municipal Code Title 17 — Zoning", "https://library.municode.com/co/broomfield/codes/municipal_code?nodeId=TIT17ZO", "Complete Broomfield zoning code with district standards and uses"),
    (BC, BM, "Planning & Zoning", "Building Division & Permits", "https://www.broomfield.org/174/Building", "Building permits, inspections, and code compliance"),
    (BC, BM, "Planning & Zoning", "Public Access Permit Search", "https://www.broomfield.org/2165/Public-Access-Permit-Search", "Online permit and development project search"),
    # Floodplain / FEMA
    (BC, BM, "Floodplain / FEMA", "Broomfield Floodplain Information", "https://www.broomfield.org/4101/Floodplains", "Local FIRM updates and flood risk areas in Broomfield"),
    (BC, BM, "Floodplain / FEMA", "FEMA Map Service Center", "https://msc.fema.gov/", "Official FEMA FIRM maps by address or community"),
    (BC, BM, "Floodplain / FEMA", "Colorado Hazard Mapping Portal", "https://coloradohazardmapping.com/hazardMapping/floodplainMapping", "Statewide floodplain and hazard mapping resource"),
    # Economic Development
    (BC, BM, "Economic Development", "Economic Vitality Department", "https://www.broomfield.org/247/Economic-Vitality", "Commercial and industrial development support for Broomfield"),
    (BC, BM, "Economic Development", "Business Assistance Resources", "https://broomfield.org/3970/Business-Assistance-Resources", "Finance, technical assistance, and development opportunities"),
    # Schools
    (BC, BM, "Schools", "Adams 12 Five Star Schools", "https://www.adams12.org/", "Primary district serving most of Broomfield (54 schools)"),
    (BC, BM, "Schools", "Boulder Valley School District", "https://www.bvsd.org/", "Serves portions of Broomfield near Boulder County border"),
    # DOLA / State
    (BC, "REGIONAL / ALL MUNICIPALITIES", "Property Assessment", "DOLA — Division of Property Taxation", "https://dpt.colorado.gov/locality/broomfield-county-assessor", "State DPT page for Broomfield County Assessor"),
]

for row_data in broomfield_rows:
    append_row(*row_data)

# ── WELD COUNTY ──────────────────────────────────────────────────────────────
WC = "Weld County"
WU = "Weld County (Unincorporated)"

weld_rows = [
    # Property Assessment
    (WC, WU, "Property Assessment", "Weld County Assessor Office", "https://www.weld.gov/Government/Departments/Assessor", "Official assessor office — property discovery, valuation, and classification"),
    (WC, WU, "Property Assessment", "Property Data Search Portal", "https://www.co.weld.co.us/apps1/propertyportal/", "Search by account number, name, address, subdivision, or parcel"),
    (WC, WU, "Property Assessment", "Property Portal Map Search", "https://maps.weld.gov/propertyportal/", "Interactive GIS map for parcel and property data search"),
    (WC, WU, "Property Assessment", "NOV & Appeals FAQ", "https://www.weld.gov/Government/Departments/Assessor/FAQs/Frequently-Asked-Questions", "Notice of Valuation appeal info and assessment procedures"),
    (WC, WU, "Property Assessment", "Board of Equalization", "https://www.weld.gov/Government/Departments/Commissioners/County-Board-of-Equalization", "BOE valuation appeals (September 1 – November 1)"),
    # Property Tax / Treasurer
    (WC, WU, "Property Tax / Treasurer", "Weld County Tax Portal", "https://www.weldtax.com/treasurer/web/login.jsp", "Online tax account search and property tax payments"),
    (WC, WU, "Property Tax / Treasurer", "Pay Taxes Online", "https://www.weld.gov/Info/Pay-Your-Taxes-Online", "Online property tax payment system"),
    (WC, WU, "Property Tax / Treasurer", "Mill Levy Report", "https://www.weld.gov/Government/Departments/Assessor/Taxing-Authority-Information/Mill-Levy-Report", "Mill levy for 150 tax districts and 1,500+ tax areas"),
    (WC, WU, "Property Tax / Treasurer", "Taxing Authority Information", "https://www.weld.gov/Government/Departments/Assessor/Taxing-Authority-Information", "All taxing authorities and mill levy data by district"),
    (WC, WU, "Property Tax / Treasurer", "Treasurer FAQ", "https://www.weld.gov/Government/Departments/Treasurer-Public-Trustee/Treasurer/FAQ", "Property tax and treasurer services FAQ"),
    # Clerk & Recorder
    (WC, WU, "Clerk & Recorder / Official Records", "Weld County Clerk & Recorder", "https://www.weld.gov/Government/Departments/Clerk-and-Recorder", "Official records — deeds and public documents since 1865"),
    (WC, WU, "Clerk & Recorder / Official Records", "Recording Department", "https://www.weld.gov/Government/Departments/Clerk-and-Recorder/Recording-Department", "Document recording and preservation services"),
    (WC, WU, "Clerk & Recorder / Official Records", "Recorded Documents Portal", "https://www.weld.gov/Government/County-Data/Recorded-Documents", "Self-service portal for accessing recorded documents"),
    # GIS / Mapping
    (WC, WU, "GIS / Mapping", "Weld County GIS Hub", "https://gishub.weldgov.com/", "GIS data portal — parcels, zoning, aerial imagery; weekly-updated datasets"),
    (WC, WU, "GIS / Mapping", "GIS Data Download", "https://gishub.weldgov.com/pages/data", "Downloadable parcel, zoning, floodplain, and spatial datasets"),
    (WC, WU, "GIS / Mapping", "Interactive Maps", "https://gishub.weldgov.com/pages/interactive-maps", "Collection of web maps for planning, zoning, and property data"),
    # Planning & Zoning
    (WC, WU, "Planning & Zoning", "Planning and Development Services", "https://www.weld.gov/Government/Departments/Planning-and-Zoning", "County planning, zoning, comprehensive plans, and land use regulations"),
    (WC, WU, "Planning & Zoning", "Comprehensive Plan (Ch. 22)", "https://library.municode.com/co/weld_county/codes/charter_and_county_code?nodeId=CH22COPL", "Weld County Comprehensive Plan — policy for land use decisions"),
    (WC, WU, "Planning & Zoning", "Zoning Ordinance (Ch. 23)", "https://library.municode.com/co/weld_county/codes/charter_and_county_code?nodeId=CH23ZO", "Zoning districts, overlays, and land use regulations"),
    (WC, WU, "Planning & Zoning", "Current Planning Applications", "https://www.weld.gov/Government/Departments/Planning-and-Zoning/Current-Planning", "Land use application review and permit processing"),
    # Floodplain / FEMA
    (WC, WU, "Floodplain / FEMA", "Weld County Floodplain Management", "https://www.weld.gov/Government/Departments/Planning-and-Zoning/Floodplain-Management", "County floodplain regulations and flood hazard management"),
    (WC, WU, "Floodplain / FEMA", "Colorado Hazard Mapping — Weld County", "https://coloradohazardmapping.com/county/Weld", "Flood hazard and erosion data specific to Weld County"),
    (WC, WU, "Floodplain / FEMA", "FEMA Map Service Center", "https://msc.fema.gov/", "Official FEMA FIRM maps by address"),
    (WC, WU, "Floodplain / FEMA", "Colorado Fluvial Hazard Zone Mapping", "https://www.coloradofhz.com/", "State erosion, sediment, and dynamic river hazard mapping"),
    # Economic Development
    (WC, WU, "Economic Development", "Upstate Colorado Economic Development", "https://upstatecolorado.org/", "Public/private non-profit for Weld County business investment"),
    (WC, WU, "Economic Development", "Metro Denver Economic Council — Weld", "https://www.metrodenver.org/do-business/communities/weld", "Regional economic data and development info for Weld County"),
    # Schools
    (WC, WU, "Schools", "Greeley-Evans School District 6", "https://www.greeleyschools.org/", "Primary district serving Greeley, Evans, and unincorporated Weld County"),
    (WC, WU, "Schools", "Windsor RE-4 School District", "https://www.weldre4.org/", "School district serving Windsor and surrounding areas"),
    (WC, WU, "Schools", "Weld RE-5J School District", "https://www.weldre5j.org/", "School district serving northeast Weld County"),
    (WC, WU, "Schools", "Weld RE-1 School District (Platteville)", "https://wcsdre1.org/en-US", "School district serving Platteville and surrounding areas"),
    # DOLA / State
    (WC, WU, "Property Assessment", "DOLA — Division of Property Taxation", "https://dpt.colorado.gov/locality/weld-county-assessor", "State DPT page for Weld County Assessor"),

    # Greeley
    (WC, "Greeley", "Planning & Zoning", "City of Greeley Planning & Zoning", "https://greeleyco.gov/business/construction-and-growth/planning-and-zoning/", "Development review, zoning compliance, and planning applications"),
    (WC, "Greeley", "Planning & Zoning", "Greeley Planning Applications & Permits", "https://greeleyco.gov/business/construction-and-growth/planning-and-zoning/planning-applications-and-permits/", "Land use applications, site plans, and development permits"),
    (WC, "Greeley", "Planning & Zoning", "Greeley Municipal Development Code", "https://library.municode.com/co/greeley/codes/municipal_code", "Title 24 development code effective October 1, 2021"),

    # Windsor
    (WC, "Windsor", "Planning & Zoning", "Town of Windsor Planning & Zoning", "https://www.windsorgov.com/1229/Planning-Zoning", "Zoning administration, development review, and annexation proposals"),
    (WC, "Windsor", "Planning & Zoning", "Windsor Development Center", "https://windsorgov.com/949/Development-Center", "Development review process and application assistance"),
    (WC, "Windsor", "Planning & Zoning", "Windsor Zoning Code", "https://library.municode.com/co/windsor/codes/charter_and_municipal_code?nodeId=WI_CH16ZO", "Windsor municipal zoning ordinance and land use regulations"),

    # Evans
    (WC, "Evans", "Planning & Zoning", "City of Evans Planning & Development", "https://www.evanscolorado.gov/building-business-development/planning-and-development/", "Project review, approvals, and zoning administration"),
    (WC, "Evans", "Planning & Zoning", "Evans Land Development Code", "https://library.municode.com/co/evans/codes/municipal_code?nodeId=MUCO_TIT18LADECO", "Development standards and land use regulations"),
    (WC, "Evans", "Planning & Zoning", "Evans Permits & Licensing", "https://www.evanscolorado.gov/your-government/applications-licenses-and-permits/", "Permit applications and land development procedures"),

    # Fort Lupton
    (WC, "Fort Lupton", "Planning & Zoning", "City of Fort Lupton Planning & Building", "https://www.fortluptonco.gov/159/Planning-Building", "Development review and permitting"),
    (WC, "Fort Lupton", "Planning & Zoning", "Fort Lupton Development Code", "https://www.fortluptonco.gov/856/Fort-Lupton-Development-Code", "Full land use and zoning regulations"),
    (WC, "Fort Lupton", "Planning & Zoning", "Fort Lupton Zoning Verification", "https://www.fortluptonco.gov/405/Zoning", "Zoning district information and verification requests"),

    # Firestone
    (WC, "Firestone", "Planning & Zoning", "Town of Firestone Planning & Development", "https://www.firestoneco.gov/145/Planning-Development", "Annexation, zoning, plat, and site plan review"),
    (WC, "Firestone", "Planning & Zoning", "Firestone Development Code", "https://www.firestoneco.gov/175/Development-Code", "Unified development code — zoning, subdivision, and processes"),
    (WC, "Firestone", "Planning & Zoning", "Firestone Maps & Development Plans", "https://www.firestoneco.gov/513/Maps-Development-Plans", "Interactive mapping for Firestone development documents"),

    # Frederick
    (WC, "Frederick", "Planning & Zoning", "Town of Frederick Planning", "https://www.frederickco.gov/349/Planning", "Development application review and coordination"),
    (WC, "Frederick", "Planning & Zoning", "Frederick Development Center", "https://www.frederickco.gov/96/Development-Center", "Development services and application processing"),
    (WC, "Frederick", "Planning & Zoning", "Frederick Land Use Code", "https://www.frederickco.gov/1043/Land-Use-Code", "Updated land use code effective June 1, 2026"),

    # Johnstown
    (WC, "Johnstown", "Planning & Zoning", "Town of Johnstown Planning & Development", "https://johnstownco.gov/225/Planning-Development", "Land use development application management"),
    (WC, "Johnstown", "Planning & Zoning", "Johnstown Comprehensive Plan", "https://johnstownco.gov/537/Master-Plans", "Comprehensive Plan and Downtown Master Plan"),

    # Mead
    (WC, "Mead", "Planning & Zoning", "Town of Mead Planning Division", "https://www.townofmead.org/development/page/planning-division", "Annexation, zoning, plat, and site plan services"),

    # Platteville
    (WC, "Platteville", "Planning & Zoning", "Town of Platteville Planning & Zoning", "https://platteville.colorado.gov/planning-and-zoning", "Development approvals and planning administration"),
    (WC, "Platteville", "Planning & Zoning", "Platteville Municipal Code", "https://platteville.colorado.gov/municipal-code", "Development guidelines (Chapters 15–18)"),

    # Milliken
    (WC, "Milliken", "Economic Development", "Town of Milliken Economic Development", "https://www.millikenco.gov/168/Economic-Development", "Economic development and growth initiatives"),

    # Kersey
    (WC, "Kersey", "Planning & Zoning", "Town of Kersey Planning Department", "https://www.kerseygov.com/departments/planning_department/index.php", "Land development aligned with community vision"),
]

for row_data in weld_rows:
    append_row(*row_data)

print(f"Total rows after additions: {ws.max_row}")

# ── UPDATE COUNTY DROPDOWN ───────────────────────────────────────────────────
# Remove old DV and replace with 4-county list
ws.data_validations.dataValidation = [
    dv for dv in ws.data_validations.dataValidation
    if dv.sqref != "D3"
]

dv = DataValidation(
    type="list",
    formula1='"Adams County,Boulder County,Broomfield County,Weld County"',
    showDropDown=False,
    showErrorMessage=True,
    errorTitle="Invalid County",
    error="Please select a county from the list."
)
dv.sqref = "D3"
ws.add_data_validation(dv)
print("County dropdown updated to: Adams, Boulder, Broomfield, Weld")

wb.save(DST)
print("Saved.")
