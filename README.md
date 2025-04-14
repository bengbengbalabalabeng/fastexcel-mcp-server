# FastExcel MCP Server

Java server implementing Model Context Protocol (MCP) for Excel operations with standardized header-to-column relationships.

## Features

- Read Excel file (XLSX, XLS, CSV) headers and data rows.
- Validate input paths against configured workspaces.
- Support for multiple sheets and custom header rows.
- Retrieve total number of data rows (excluding headers).

**Note**: The server will only allow operations within directories specified via `env: MCP_WORKSPACES`.

## API

### Tools

- **get_total_rows_number**
    - Get the total number of data rows (excluding the header row) in an Excel file.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns the total count of data rows as an integer.

- **read_head_spec**
    - Parse and return the header information from an Excel file.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns a sorted list of json objects containing column index and header titles.

- **read_rows_spec**
    - Parse and return the data rows from an Excel file with header association.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `readRowNumbers` (integer, optional): Number of data rows to read (excludes header). If null, reads all rows.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns a list of json containing row data linked to their headers.

## Usage with Claude Desktop

Add this to your `claude_desktop_config.json`:

### Java

```json
{
  "mcpServers": {
    "fastexcel": {
      "command": "java",
      "args": [
        "-jar",
        "<YOUR_PATH>/fastexecl-mcp-server-0.0.1-SNAPSHOT.jar"
      ],
      "env": {
        "MCP_WORKSPACES": "<YOUR_MULTIPLE_WORKSPACES_SEPARATED_BY_COMMAS>"
      }
    }
  }
}
```

## Build

### Required

- JDK 17+ GraalVM
- Maven 3.9.6+

Java build:

```bash
mvn clean package -DskipTests=true
```

## License

This MCP server is licensed under the Apache License 2.0. For more details, please see the LICENSE file in the project repository.

---

If you need further customization or integration details, let me know!