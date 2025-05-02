# mysheet MCP Server

Mysheet-mcp  is a MCP Server for converting Excel files into JSON format, particularly useful for data processing when called within large models or Agents. This project supports parsing Excel files and temporarily storing the files contained within the Excel in a configured temporary directory.

## Main Features

- **Excel to JSON**: Converts uploaded Excel files into JSON format for further processing.
- **File Storage**: Files contained within the Excel are temporarily stored in the configured `tmp` directory.

## Configuration

In `application.yml`, you can configure the following parameters:

- `upload.tmp`: Specifies the directory for temporary file storage.
- `upload.path`: Specifies the directory for uploaded file storage.

## Usage

1. **Install Dependencies**: Ensure that your project has installed the necessary dependencies.
2. **Configure Paths**: Configure `upload.tmp` and `upload.path` in `application.yml`.
3. **Call the Service**: Convert Excel files to JSON by calling the `excel2Json` method in `Excel2JsonService`.

### Example Code

