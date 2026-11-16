import requests
import json
import logging
from requests.auth import HTTPBasicAuth
import urllib3
import argparse
import sys
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.chart.axis import ChartLines


# Logging setup
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

urllib3.disable_warnings()
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ServiceNow configuration
with open(".env") as f:
    for line in f:
        line = line.strip()
        if line and not line.startswith("#"):
            key, value = line.split("=", 1)
            os.environ[key] = value


snow_config_dev = {
    'instance': os.getenv("SNOW_DEV_INSTANCE") , 
    'client_id':  os.getenv("SNOW_DEV_CLIENT_ID"),
    'client_secret': os.getenv("SNOW_DEV_CLIENT_SECRET")
}


snow_config_prod = {
    'instance': os.getenv("SNOW_PROD_INSTANCE"),
    'client_id': os.getenv("SNOW_PROD_CLIENT_ID"),
    'client_secret': os.getenv("SNOW_PROD_CLIENT_SECRET")
}



# Field sets per table
SYS_USER_FIELDS = [
    'user_name', 'name', 'email', 'phone', 'mobile_phone', 'title',
    'department', 'manager', 'location', 'active', 'locked_out',
    'last_login_time', 'sys_id', 'roles'
]


CMDB_NETGEAR_FIELDS = [
    'name', 'ip_address', 'u_application_service', 'u_business_service',
    'u_criticality', 'classification', 'location', 'operational_status',
    'location.name', 'location.city', 'location.country', 'location.u_site_code',
    'sys_id', 'owned_by', 'u_zabbix_agent', 'os'
]


KB_KNOWLEDGE_FIELDS = [
    'number', 'short_description', 'text', 'workflow_state', 'kb_category',
    'author', 'published', 'valid_to', 'sys_id', 'sys_created_by',
    'sys_created_on', 'sys_updated_on'
]


SERVER_COLUMN_MAPPING = {
    "company": "Company",
    "u_business_service": "Business Service",
    "u_application_service": "Application Service",
    "name": "Name",
    "ip_address": "IP Address",
    "location": "Location",
    "owned_by": "Owned by",
    "sys_created_on": "Created",
    "sys_updated_on": "Updated",
    "sys_updated_by": "Updated by",
    "operational_status": "Operational status",
    "business_unit": "Business Unit",
    "dns_domain": "DNS Domain",
    "u_dr": "DR",
    "u_database": "Database",
    "u_decom_date": "Decom Date",
    "default_gateway": "Default Gateway",
    "department": "Department",
    "short_description": "Description",
    "dr_backup": "Disaster backup",
    "discovery_source": "Discovery source",
    "disk_space": "Disk space (GB)",
    "os_domain": "Domain",
    "u_domain_joined": "Domain Joined",
    "sys_domain_path": "Domain Path",
    "due": "Due",
    "due_in": "Due in",
    "duplicate_of": "Duplicate Of",
    "environment": "Environment",
    "host_name": "Host name",
    "install_status": "Install Status",
    "install_date": "Installed",
    "internet_facing": "Internet Facing",
    "managed_by_group": "Managed By Group",
    "managed_by": "Managed by",
    "os": "Operating System",
    "u_patching_criticality": "Patching Criticality",
    "service_tag": "Service tag",
    "subcategory": "Subcategory",
    "used_for": "Used for",
    "u_vlan_id": "VLAN ID"
}


TABLE_CONFIG = {
    "sys_user": {"fields": SYS_USER_FIELDS,"filename": "sys_user","column_mapping": None,"query": None},
    "cmdb_ci_netgear": {"fields": CMDB_NETGEAR_FIELDS,"filename": "cmdb_ci_netgear","column_mapping": None,"query": None},
    "kb_knowledge": {"fields": KB_KNOWLEDGE_FIELDS,"filename": "kb_knowledge","column_mapping": None,"query": None},
    "cmdb_ci_server": {"fields": list(SERVER_COLUMN_MAPPING.keys()),"filename": "all_servers","column_mapping": SERVER_COLUMN_MAPPING,"query": None}
}




def get_oauth_token(config):
    """Obtain OAuth token using client credentials."""    
    
    token_url = f"https://{config['instance']}.service-now.com/oauth_token.do"
    data = {'grant_type': 'client_credentials'}
    try:
        response = requests.post( token_url, data=data, auth=HTTPBasicAuth(config['client_id'], config['client_secret']), timeout=60, verify=False )
        response.raise_for_status()
        token = response.json().get('access_token')
        logger.info("Successfully obtained access token.\n")
        return token
    except requests.exceptions.RequestException as e:
        logger.error(f"Failed to obtain OAuth token: {str(e)}")
        if e.response is not None:
            logger.error(f"Response content: {e.response.text}")
        return None



def fetch_table(config, token, table, fields, filename,query=None, column_mapping=None, statistics=False):

    base_url = (f"https://{config['instance']}.service-now.com/api/now/table/{table}")

    params = {
        'sysparm_limit': '10000',
        'sysparm_display_value': 'true',
        'sysparm_exclude_reference_link': 'true'
    }

    # Request only configured fields
    if fields:
        params['sysparm_fields'] = ','.join(fields)

    if query:
        params['sysparm_query'] = query

    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json'
    }

    try:
        logger.info(f"Querying {table}...")

        response = requests.get(base_url,headers=headers,params=params,timeout=60,verify=False)
        response.raise_for_status()
        results = response.json().get('result', [])

        logger.info(f"Fetched {len(results)} records from {table}.")
        export_to_files(results,filename,column_mapping)

        # Generate additional statistics if requested.
        if statistics:
            if (table == "cmdb_ci_server"):
                generate_server_statistics(results, filename,True)
            else:            
                logger.info(f"⚠️  Statistics are currently available only for cmdb_ci_server.\n")

        return results

    except requests.exceptions.RequestException as e:
        logger.error(f"ServiceNow API error on {table}: {str(e)}")
        if e.response is not None:
            logger.error(f"Response content: {e.response.text}")
        return []




def check_access(config, token):

    base_url = (f"https://{config['instance']}.service-now.com/api/now/table/sys_user_has_role")

    params = {
        'sysparm_limit': '10000',
        'sysparm_display_value': 'true',
        'sysparm_exclude_reference_link': 'true'
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json'
    }

    print("🔍 Retrieving ServiceNow table definitions...")
    tables_to_test = [
        "incident",
        "problem",
        "change_request",
        "sc_request",
        "sc_req_item",
        "sc_task",
        "cmdb_ci",
        "cmdb_ci_server",
        "sys_user",
        "sys_user_group",
        "sys_user_has_role",
        "kb_knowledge",
        "alm_asset",
        "core_company",
        "sys_db_object"
    ]

    accessible_tables = []
    for table in tables_to_test:
        response = requests.get(f"https://{config['instance']}.service-now.com/api/now/table/{table}",headers=headers,params={"sysparm_limit": 1},verify=False)

        if response.status_code == 200:
            accessible_tables.append(table)
            print(f"✅ {table}")
        elif response.status_code == 403:
            print(f"🔒 {table}: Access denied")
        elif response.status_code == 404:
            print(f"❌ {table}: Not found")
        else:
            print(f"⚠️ {table}: HTTP {response.status_code}")

    print("\nAccessible tables:")
    print(accessible_tables)




def export_to_files(data, filename, column_mapping=None):

    # JSON export (original ServiceNow field names)
    json_output_file = filename + ".json"

    with open(json_output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    
    df = pd.json_normalize(data)
    if column_mapping:        
        df = df.reindex(columns=list(column_mapping.keys()))        
        df = df.rename(columns=column_mapping)

    excel_output_file = filename + ".xlsx"
    df.to_excel(excel_output_file,index=False,engine='openpyxl')
    logger.info(f"✅ Saved {len(data)} records to: {excel_output_file}, {json_output_file}\n")




def create_server_chart(ws, sheet_name, counts):
    """
    Create a clean Excel-style horizontal clustered bar chart.   
    """

    # --------------------------------------------------
    # 1. PREPARE TOP 10 + OTHERS
    # --------------------------------------------------
    top_counts = counts.head(10).copy()
    remaining = int(counts.iloc[10:].sum())

    if remaining > 0:
        top_counts.loc["Others"] = remaining

    # Ascending order places the largest category at the bottom.
    top_counts = top_counts.sort_values(ascending=False)

    # --------------------------------------------------
    # 2. WRITE HELPER DATA
    # --------------------------------------------------

    ws.cell( row=1, column=5,value="Category")
    ws.cell(row=1, column=6, value="Server Count")

    for row, (name, count) in enumerate(top_counts.items(),start=2):
        display_name = str(name)

        # Shorten only extremely long chart labels.
        if len(display_name) > 42:
            display_name = display_name[:39] + "..."

        ws.cell( row=row,column=5, value=display_name)
        ws.cell( row=row, column=6, value=int(count) )

    last_row = len(top_counts) + 1

    # --------------------------------------------------
    # 3. CREATE CHART
    # --------------------------------------------------
    chart = BarChart()

    chart.type = "bar"
    chart.grouping = "clustered"
    chart.overlap = 0
    chart.gapWidth = 70

    chart.title = f"Servers per {sheet_name}"

    chart.height = 10
    chart.width = 18

    chart.legend = None

    # No redundant axis titles.
    chart.x_axis.title = None
    chart.y_axis.title = None

    # --------------------------------------------------
    # 4. ADD DATA
    # --------------------------------------------------
    data = Reference(ws, min_col=6, max_col=6, min_row=2, max_row=last_row )
    categories = Reference( ws, min_col=5, max_col=5, min_row=2, max_row=last_row )
    chart.add_data( data, titles_from_data=False, from_rows=False)
    chart.set_categories(categories)

    # --------------------------------------------------
    # 5. FORMAT BARS
    # --------------------------------------------------
    series = chart.series[0]

    # Prevent hollow/white bars.
    series.invertIfNegative = False

    bar_color = "4472C4"

    series.graphicalProperties.noFill = False
    series.graphicalProperties.solidFill = bar_color

    series.graphicalProperties.line.noFill = False
    series.graphicalProperties.line.solidFill = bar_color

    # Remove individual data-point formatting.
    series.dPt = []

    # --------------------------------------------------
    # 6. CATEGORY AXIS
    # --------------------------------------------------
    chart.x_axis.delete = False
    chart.x_axis.axPos = "l"
    chart.x_axis.tickLblPos = "nextTo"

    chart.x_axis.tickLblSkip = 1
    chart.x_axis.lblOffset = 100

    chart.x_axis.majorTickMark = "none"
    chart.x_axis.minorTickMark = "none"

    chart.x_axis.scaling.orientation = "maxMin"
    
    chart.x_axis.crosses = "autoZero"

    # --------------------------------------------------
    # 7. CALCULATE NUMERICAL INTERVAL
    # --------------------------------------------------
    maximum_value = int(top_counts.max())

    if maximum_value <= 10:
        major_unit = 1
    elif maximum_value <= 20:
        major_unit = 5
    elif maximum_value <= 50:
        major_unit = 10
    elif maximum_value <= 100:
        major_unit = 20
    elif maximum_value <= 250:
        major_unit = 50
    elif maximum_value <= 500:
        major_unit = 100
    elif maximum_value <= 1000:
        major_unit = 200
    elif maximum_value <= 2500:
        major_unit = 200
    elif maximum_value <= 5000:
        major_unit = 500
    else:
        major_unit = 1000

    # --------------------------------------------------
    # 8. NUMERICAL AXIS
    # --------------------------------------------------
    chart.y_axis.delete = False

    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.orientation = "minMax"

    chart.y_axis.majorUnit = major_unit
    chart.y_axis.numFmt = "0"

    # Position the numerical axis at the bottom.
    #chart.y_axis.axPos = "b"
    chart.y_axis.tickLblPos = "nextTo"

    chart.y_axis.majorTickMark = "out"
    chart.y_axis.minorTickMark = "none"

    # Add vertical gridlines.
    chart.y_axis.majorGridlines = ChartLines()

    chart.dataLabels = None

    return chart



def generate_server_statistics(data, filename,onlyOperational=True):

    logger.info("Generating server statistics...")

    df = pd.json_normalize(data)
    if onlyOperational:
        df = df[df["operational_status"] == "Operational"]

    statistics_fields = {
        "company": "Companies",
        "u_business_service": "Business Services",
        "location": "Locations",
        "owned_by": "Owned by",
        "os": "Operating Systems"
    }

    # --------------------------------------------------
    # 2. CREATE WORKBOOK
    # --------------------------------------------------
    wb = Workbook()

    dashboard = wb.active
    dashboard.title = "General"

    header_fill = PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    header_font = Font(color="FFFFFF",bold=True)
    title_font = Font(size=18,bold=True,color="1F4E78")

    
    # --------------------------------------------------
    # 3. DASHBOARD HEADER
    # --------------------------------------------------
    dashboard["A1"] = "ServiceNow CMDB Server Statistics"
    dashboard["A1"].font = title_font

    dashboard["A3"] = "Infrastructure Overview"
    dashboard["A3"].font = Font(bold=True, size=12)

    dashboard["A4"] = "Total Servers"
    dashboard["B4"] = len(df)

    dashboard["A5"] = "Statistics generated"
    dashboard["B5"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

    summary_row = 7

    # --------------------------------------------------
    # 4. PROCESS STATISTICS FIELDS
    # --------------------------------------------------
    for field, sheet_name in statistics_fields.items():

        if field not in df.columns:
            logger.warning(f"Field '{field}' not found in exported data.")
            continue

        # Normalize missing values.
        values = (df[field].fillna("Not specified").astype(str).str.strip())

        # Treat empty strings and common empty placeholders
        # as missing values.
        values = values.replace(["", "-", "None", "null", "nan"],"Not specified")

        # Count occurrences.
        counts = (values.value_counts().sort_values(ascending=False))

        # --------------------------------------------------
        # 5. DASHBOARD SUMMARY
        # --------------------------------------------------

        dashboard.cell(row=summary_row,column=1,value=f"Unique {sheet_name}")
        dashboard.cell(row=summary_row,column=2,value=len(counts))

        summary_row += 1

        # --------------------------------------------------
        # 6. CREATE STATISTICS WORKSHEET
        # --------------------------------------------------

        ws = wb.create_sheet(sheet_name)
        ws.append([sheet_name,"Server Count","Percentage"])

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

            cell.alignment = Alignment(horizontal="center")

        # Populate complete statistics table.
        for value, count in counts.items():

            percentage = (count / len(df) if len(df) > 0  else 0)
            ws.append([value,int(count),percentage])

        # Format percentage column.
        for cell in ws["C"][1:]:
            cell.number_format = "0.00%"

        # Worksheet formatting.
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18

        ws.column_dimensions["E"].width = 35
        ws.column_dimensions["F"].width = 18

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:C{len(counts) + 1}"

        # --------------------------------------------------
        # 7. CREATE CHARTS
        # --------------------------------------------------
        if len(counts) > 0:
            # Individual worksheet chart.
            worksheet_chart = create_server_chart(ws,sheet_name,counts)
            ws.add_chart(worksheet_chart, "H2")
    

    # --------------------------------------------------
    # 8. DASHBOARD FORMATTING
    # --------------------------------------------------
    dashboard.column_dimensions["A"].width = 30
    dashboard.column_dimensions["B"].width = 22

    dashboard.freeze_panes = "A7"

    for cell in dashboard[3]:
        cell.font = Font(bold=True,size=12)

    # --------------------------------------------------
    # 9. SAVE WORKBOOK
    # --------------------------------------------------
    statistics_filename = f"{filename}_statistics.xlsx"
    wb.save(statistics_filename)

    logger.info(f"Statistics workbook successfully generated: {statistics_filename}")
    logger.info(f"Processed {len(df)} servers.")




def main():

    parser = argparse.ArgumentParser(description="ServiceNow table export utility",epilog="""
Examples:
  python .\\ServiceNowExporter.py --table cmdb_ci_server 
  python .\\ServiceNowExporter.py --table cmdb_ci_server --outfile All_servers.xlsx
  python .\\ServiceNowExporter.py --table cmdb_ci_server --query "ip_address=10.10.10.10"
  python .\\ServiceNowExporter.py --table cmdb_ci_server --query "company.name=Contoso" 
  python .\\ServiceNowExporter.py --table cmdb_ci_server --query "company.name=Contoso^name=CNTSPRODHANA1^u_business_service=Contoso HANA Production"
  python .\\ServiceNowExporter.py --outfile custom_output
  python .\\ServiceNowExporter.py --check 
  python .\\ServiceNowExporter.py --table cmdb_ci_server --statistics

""",
    formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--table",required=False,choices=list(TABLE_CONFIG.keys()) + ["all"])
    parser.add_argument("--query",help="Search for specific entries on cmdb fields, e.g name=CNTSPRODHANA1")
    parser.add_argument("--outfile",help="Override the default output file")
    parser.add_argument("--env",required=False,help="The environment to access: dev/prod",choices=["dev","prod"],default="prod")
    parser.add_argument("--check",required=False, action='store_true',help="Check the current user's access to various tables")
    parser.add_argument("--statistics",action="store_true",help="Generate additional statistics. Currently only for cmdb_ci_server")
    args = parser.parse_args()

    print(f"\n✅ Running in {args.env} environment\n")
    snow_config = snow_config_dev if args.env == "dev" else snow_config_prod
    access_token = get_oauth_token(snow_config)

    if not access_token:
        logger.error("Could not retrieve OAuth access token.")
        sys.exit(1)


    if (args.check):
        check_access(config=snow_config,token=access_token)
        sys.exit(1)

    if (not args.table):
        logger.error("Please supply a table value to query.")
        sys.exit(1)

    tables = (list(TABLE_CONFIG.keys()) if args.table == "all" else [args.table] )
    for table_name in tables:

        logger.info(f"Processing table: {table_name}")
        table_settings = TABLE_CONFIG[table_name]
        
        export_file = args.outfile if (args.outfile) else table_settings["filename"]
        query_value = args.query if (args.query) else table_settings["query"]
        fetch_table(config=snow_config,token=access_token,table=table_name,fields=table_settings["fields"],filename=export_file,
                    query=query_value,column_mapping=table_settings["column_mapping"],statistics=args.statistics
        )


if __name__ == "__main__":
    main()