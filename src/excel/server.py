# from mcp.server.fastmcp import FastMCP, Context
from fastmcp import FastMCP
import pandas as pd
import os
from typing import List, Dict, Optional
import json

# Create the MCP server instance with required dependencies
mcp = FastMCP("Excel Manager", dependencies=["pandas", "openpyxl", "fastmcp"])

# Track open workbooks to maintain state
WORKBOOKS: Dict[str, pd.ExcelFile] = {}
DATAFRAMES: Dict[str, Dict[str, pd.DataFrame]] = {}

def normalize_path(path: str) -> str:
    """Convert file path to absolute path"""
    return os.path.abspath(os.path.expanduser(path))

@mcp.resource("excel://{file_path}/sheets")
def list_sheets(file_path: str) -> str:
    """List all sheets in an Excel file"""
    path = normalize_path(file_path)
    if path not in WORKBOOKS:
        WORKBOOKS[path] = pd.ExcelFile(path)
        DATAFRAMES[path] = {}
    
    sheets = WORKBOOKS[path].sheet_names
    return json.dumps({
        "file": path,
        "sheets": sheets
    })

@mcp.resource("excel://{file_path}/sheet/{sheet_name}")
def get_sheet_data(file_path: str, sheet_name: str) -> str:
    """Get the contents of a specific sheet as JSON"""
    path = normalize_path(file_path)
    if path not in DATAFRAMES or sheet_name not in DATAFRAMES[path]:
        if path not in WORKBOOKS:
            WORKBOOKS[path] = pd.ExcelFile(path)
            DATAFRAMES[path] = {}
        DATAFRAMES[path][sheet_name] = pd.read_excel(WORKBOOKS[path], sheet_name)
    
    df = DATAFRAMES[path][sheet_name]
    return df.to_json(orient='records', date_format='iso')

@mcp.tool()
def read_excel(file_path: str, sheet_name: Optional[str] = None) -> str:
    """
    Read an Excel file and return its contents.
    If sheet_name is provided, reads that specific sheet, otherwise reads all sheets.
    """
    path = normalize_path(file_path)
    try:
        if sheet_name:
            df = pd.read_excel(path, sheet_name=sheet_name)
            return df.to_string()
        else:
            dfs = pd.read_excel(path, sheet_name=None)
            return "\n\n".join(f"Sheet: {name}\n{df.to_string()}" for name, df in dfs.items())
    except Exception as e:
        return f"Error reading Excel file: {str(e)}"

@mcp.tool()
def query_excel(file_path: str, sheet_name: str, query: str) -> str:
    """
    Run a query on Excel data using pandas query syntax.
    Example query: "age > 30 and department == 'Sales'"
    """
    path = normalize_path(file_path)
    try:
        if path not in DATAFRAMES or sheet_name not in DATAFRAMES[path]:
            if path not in WORKBOOKS:
                WORKBOOKS[path] = pd.ExcelFile(path)
                DATAFRAMES[path] = {}
            DATAFRAMES[path][sheet_name] = pd.read_excel(WORKBOOKS[path], sheet_name)
        
        df = DATAFRAMES[path][sheet_name]
        result = df.query(query)
        return result.to_string()
    except Exception as e:
        return f"Error querying Excel data: {str(e)}"

@mcp.tool()
def update_cell(file_path: str, sheet_name: str, row: int, column: str, value: str) -> str:
    """
    Update a specific cell in an Excel file.
    Column should be Excel-style (A, B, C, etc.)
    """
    path = normalize_path(file_path)
    try:
        # Load the sheet if not already loaded
        if path not in DATAFRAMES or sheet_name not in DATAFRAMES[path]:
            if path not in WORKBOOKS:
                WORKBOOKS[path] = pd.ExcelFile(path)
                DATAFRAMES[path] = {}
            DATAFRAMES[path][sheet_name] = pd.read_excel(WORKBOOKS[path], sheet_name)
        
        df = DATAFRAMES[path][sheet_name]
        
        # Convert Excel column letter to numeric index
        col_idx = 0
        for char in column.upper():
            col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
        col_idx -= 1
        
        # Update the value
        df.iloc[row-1, col_idx] = value
        
        # Save the changes
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return f"Updated cell {column}{row} to '{value}' in sheet '{sheet_name}'"
    except Exception as e:
        return f"Error updating cell: {str(e)}"

@mcp.tool()
def add_row(file_path: str, sheet_name: str, values: List[str]) -> str:
    """Add a new row to the specified sheet"""
    path = normalize_path(file_path)
    try:
        if path not in DATAFRAMES or sheet_name not in DATAFRAMES[path]:
            if path not in WORKBOOKS:
                WORKBOOKS[path] = pd.ExcelFile(path)
                DATAFRAMES[path] = {}
            DATAFRAMES[path][sheet_name] = pd.read_excel(WORKBOOKS[path], sheet_name)
        
        df = DATAFRAMES[path][sheet_name]
        df.loc[len(df)] = values
        
        # Save the changes
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return f"Added new row to sheet '{sheet_name}'"
    except Exception as e:
        return f"Error adding row: {str(e)}"

@mcp.tool()
def create_sheet(file_path: str, sheet_name: str, headers: List[str]) -> str:
    """Create a new sheet in the Excel file with specified headers"""
    path = normalize_path(file_path)
    try:
        df = pd.DataFrame(columns=headers)
        
        # Save the new sheet
        with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Update our cached data
        if path in WORKBOOKS:
            del WORKBOOKS[path]
        if path in DATAFRAMES:
            del DATAFRAMES[path]
        
        return f"Created new sheet '{sheet_name}' with headers: {', '.join(headers)}"
    except Exception as e:
        return f"Error creating sheet: {str(e)}"

if __name__ == "__main__":
    mcp.run()