# FastExcel MCP Server

Java server implementing Model Context Protocol (MCP) for Excel operations with standardized header-to-column relationships.

## Features

- Read Excel file (XLSX, XLS, CSV) headers and data rows.
- Validate input paths against configured workspaces.
- Support for multiple sheets and custom header rows.
- Retrieve total number of data rows (excluding headers).
- Built-in performance caching for faster repeated access.

**Note**: The server will only allow operations within directories specified via `env: MCP_WORKSPACES`.

## API

### Tools

- **get_total_rows_number (with cache support)**
    - Get the total number of data rows (excluding the header row) in an Excel file.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns the total count of data rows as an integer.

- **get_sheet_names (with cache support)**
    - Retrieve all sheet names and their indices from an Excel file.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
    - Returns a list of json objects containing sheet number and sheet name.

- **read_head_spec (with cache support)**
    - Parse and return the header information from an Excel file.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns a sorted list of json objects containing column index and header titles.

- **read_rows_spec (with cache support)**
    - Parse and return the data rows from an Excel file with header association.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
        - `headRowNumber` (integer): Row number (1-based index) where the header is located.
        - `readRowNumbers` (integer, optional): Number of data rows to read (excludes header). If null, reads all rows.
        - `sheetName` (string, optional): Name of the sheet to read. Defaults to the first sheet if empty.
    - Returns a list of json containing row data linked to their headers.

- **cache_clear**
    - Clear all cached Excel file data from memory.
    - Returns true when cache is cleared successfully.

- **test_cache_available**
    - Test if the specified Excel file is available in cache.
    - Inputs:
        - `excelPath` (string): Absolute or relative path to the Excel file.
    - Returns a json object indicating cache status and file MD5 hash:
        - `cached`: boolean indicating if file is cached
        - `result`: string containing file MD5 hash information

## Usage with Claude Desktop

Add this to your `claude_desktop_config.json`:

### Java

```json
{
  "mcpServers": {
    "fastexcel-mcp-server": {
      "command": "java",
      "args": [
        "-jar",
        "<YOUR_PATH>/fastexcel-mcp-server-0.0.1-SNAPSHOT.jar"
      ],
      "env": {
        "MCP_WORKSPACES": "<YOUR_MULTIPLE_WORKSPACES_SEPARATED_BY_COMMAS>",
        "CACHE_INITIAL_CAPACITY": "[OPTIONAL] <MINIMUM_TOTAL_SIZE_FOR_THE_INTERNAL_DATA_STRUCTURES> <DEFAULT: 100>",
        "CACHE_MAXIMUM_SIZE": "[OPTIONAL] <MAXIMUM_NUMBER_OF_ENTRIES_THE_CACHE_MAY_CONTAIN> <DEFAULT: 1000>",
        "CACHE_EXPIRE_AFTER_WRITE": "[OPTIONAL] <LENGTH_OF_TIME_AFTER_AN_ENTRY_IS_CREATED_THAT_IT_SHOULD_BE_AUTOMATICALLY_REMOVED> <DEFAULT: 35s>"
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