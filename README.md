# ServiceNow Exporter

A Python utility for exporting ServiceNow table data to JSON and Excel, with additional statistics and charts for the `cmdb_ci_server` table.

The tool uses ServiceNow OAuth client credentials, supports development and production instances, accepts encoded ServiceNow queries, and can generate an Excel dashboard focused on operational servers.

## Features

- Export supported ServiceNow tables to JSON and Excel.
- Authenticate with OAuth 2.0 client credentials.
- Select development or production ServiceNow environments.
- Apply ServiceNow encoded queries through `sysparm_query`.
- Check API access to commonly used ServiceNow tables.
- Export basic data analytics with charts for Companies, Business services, Locations, Owners, Operating systems



## Supported tables

The current script includes configurations for:

| ServiceNow table | Default output name | Custom column mapping |
|---|---|---|
| `sys_user` | `sys_user` | No |
| `cmdb_ci_netgear` | `cmdb_ci_netgear` | No |
| `kb_knowledge` | `kb_knowledge` | No |
| `cmdb_ci_server` | `all_servers` | Yes |

Additional tables can be added in the `TABLE_CONFIG` dictionary.

## Requirements

- Python 3.10 or later recommended
- Network access to the target ServiceNow instance
- A ServiceNow OAuth application configured for the client-credentials grant
- Read access to the ServiceNow tables and fields being queried

Python dependencies:

```text
requests
urllib3
pandas
openpyxl
```

Install them with:

```bash
python -m pip install requests urllib3 pandas openpyxl
```

## Configuration

Create a `.env` file in the same directory as the script:

```dotenv
SNOW_DEV_INSTANCE=dev-instance-name
SNOW_DEV_CLIENT_ID=dev-client-id
SNOW_DEV_CLIENT_SECRET=dev-client-secret

SNOW_PROD_INSTANCE=prod-instance-name
SNOW_PROD_CLIENT_ID=prod-client-id
SNOW_PROD_CLIENT_SECRET=prod-client-secret
```

For example, if the ServiceNow URL is:

```text
https://contoso.service-now.com
```

set the instance value to:

```dotenv
SNOW_PROD_INSTANCE=contoso
```

Do not include `https://` or `.service-now.com` in the instance value.



## Usage

Display help:

```bash
python ServiceNowExporter.py --help
```

Export all configured fields from `cmdb_ci_server` using the production environment:

```bash
python ServiceNowExporter.py --table cmdb_ci_server
```

Use the development environment:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --env dev
```

Generate the server statistics workbook:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --statistics
```

Use a custom output filename:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --outfile custom_servers
```

The script appends the appropriate extension, so prefer a base filename without `.xlsx`.

Export all configured tables:

```bash
python ServiceNowExporter.py --table all
```

Check access to common ServiceNow tables:

```bash
python ServiceNowExporter.py --check
```

## ServiceNow query examples

Filter by IP address:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --query "ip_address=10.10.10.10"
```

Filter by company display value:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --query "company.name=Contoso"
```

Combine conditions with the ServiceNow encoded-query separator `^`:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --query "company.name=Contoso^name=CNTSPRODHANA1^u_business_service=Contoso HANA Production"
```

Generate statistics for a filtered server population:

```bash
python ServiceNowExporter.py --table cmdb_ci_server --query "company.name=Contoso" --statistics
```

The API query is applied first. The statistics function then keeps only records where:

```text
operational_status = Operation
```

## Output files

For a default `cmdb_ci_server` export, the script produces:

```text
all_servers.json
all_servers.xlsx
```

When `--statistics` is specified, it also produces:

```text
all_servers_statistics.xlsx
```

If `--outfile contoso_servers` is used, the output names are based on `contoso_servers`.

## Statistics workbook

The statistics workbook contains a `General` worksheet and one worksheet per configured statistics field.

### General

The sheet includes:

- Total number of operational servers included in the statistics
- Statistics generation timestamp
- Number of unique companies
- Number of unique business services
- Number of unique locations
- Number of unique owners
- Number of unique operating systems
- Embedded charts for each statistics category

### Statistics worksheets

Each statistics worksheet includes:

- Complete category list
- Server count
- Percentage of the filtered server population
- Top 10 chart categories
- An `Others` category containing the sum of remaining values

Missing, blank, or placeholder values are normalized to `Not specified`.



## Suggested future enhancements

- Add command-line options for statistics status and Top-N size.
- Add a CMDB data-completeness scorecard.
- Report internet-facing servers by owner, operating system, and business service.
- Add stale-record analysis using `sys_updated_on`.
- Add operating-system lifecycle mappings for supported, extended-support, and end-of-support platforms.
- Add ownership and management coverage metrics.
- Compare CMDB inventory with EDR, vulnerability-management, SIEM, cloud, and monitoring inventories.
- Implement pagination for ServiceNow result sets larger than 10,000 records.
- Enable TLS certificate validation with a configurable CA bundle.
- Replace local `.env` secrets with Azure Key Vault, HashiCorp Vault, GitHub Actions secrets, or another enterprise secrets manager.
- Add unit tests for filtering, normalization, Top 10 plus `Others`, and chart generation.

## Extending the table configuration

Add a field list and a new entry to `TABLE_CONFIG`:

```python
MY_TABLE_FIELDS = [
    "field_one",
    "field_two",
    "sys_id"
]

TABLE_CONFIG["my_table"] = {
    "fields": MY_TABLE_FIELDS,
    "filename": "my_table_export",
    "column_mapping": None,
    "query": None
}
```

The table automatically becomes available through the `--table` argument because the allowed choices are generated from `TABLE_CONFIG`.

## Known limitations

- `sysparm_limit` is currently set to 10,000 and pagination is not yet implemented.
- TLS certificate verification is disabled in the current implementation.
- The statistics workflow is currently specific to `cmdb_ci_server`.
- Statistics depend on ServiceNow display values because the API request uses `sysparm_display_value=true`.
- The operational-status filter expects the display value `Operation`. Confirm the exact display value used by the target ServiceNow instance.
- Chart appearance may vary slightly across Microsoft Excel versions and other spreadsheet applications.



## Contributing

1. Create a feature branch.
2. Keep API, export, statistics, and chart logic in separate functions.
3. Do not commit ServiceNow credentials or generated CMDB exports.
4. Add or update tests for behavioral changes.
5. Document new command-line parameters and supported tables.
6. Submit a pull request describing the change and its expected output.

## License

No license is included by default. Before publishing the repository, select a license appropriate for the intended use. For internal-only corporate use, confirm the organization's source-code and information-classification requirements before adding a public open-source license.

## Disclaimer

This utility exports potentially sensitive infrastructure metadata. Users are responsible for applying appropriate access controls, retention rules, information classification, and secure handling requirements to both the source data and generated output files.
