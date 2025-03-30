"""
Excel Schema - Module for handling Excel sheet structure as a schema.
"""

import json
from pathlib import Path
from typing import Dict, Any, List, Optional, Union, Sequence, cast, Mapping, TypedDict, MutableMapping

import jsonschema


# Define typed dictionaries for better type checking
class HeaderInfo(TypedDict, total=False):
    name: str
    column_letter: str
    data_type: str


class SheetStructure(TypedDict, total=False):
    headers: Dict[Union[int, str], HeaderInfo]
    columns: Dict[str, Any]
    rows: int
    columns_count: int
    merged_cells: List[str]


class ExcelStructure(TypedDict, total=False):
    sheets: Dict[str, SheetStructure]
    named_ranges: Dict[str, str]
    file_properties: Dict[str, Any]


class ExcelSchema:
    """
    Class for handling Excel sheet structures as schemas.
    Provides functionality for validation and schema operations.
    """

    def __init__(self, structure: Optional[ExcelStructure] = None):
        """
        Initialize the ExcelSchema object.

        Args:
            structure: Optional dictionary with Excel structure
        """
        default_structure: ExcelStructure = {"sheets": {}, "named_ranges": {}, "file_properties": {}}
        self.structure = structure if structure is not None else default_structure
        
    def load_structure(self, structure: ExcelStructure) -> None:
        """
        Load a structure into the schema.
        
        Args:
            structure: Dictionary with Excel structure
        """
        self.structure = structure
    
    def save_structure(self, output_file: Union[str, Path]) -> None:
        """
        Save the structure to a file.
        
        Args:
            output_file: Path to save the structure file
        """
        output_path = Path(output_file)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.structure, f, indent=2)
    
    def to_json_schema(self) -> Dict[str, Any]:
        """
        Convert the Excel structure to a JSON Schema.
        
        Returns:
            Dictionary with JSON Schema representation
        """
        schema: Dict[str, Any] = {
            "$schema": "http://json-schema.org/draft-07/schema#",
            "title": "Excel Data Schema",
            "description": f"Schema for Excel file {self.structure.get('file_properties', {}).get('filename', 'unknown')}",
            "type": "object",
            "properties": {
                "data": {
                    "type": "object",
                    "properties": {}
                }
            }
        }
        
        # Get sheets safely
        sheets = self.structure.get("sheets", {})
        data_props = schema["properties"]["data"]["properties"]
        
        # Create schema for each sheet
        for sheet_name, sheet_structure in sheets.items():
            # Get headers safely
            headers = sheet_structure.get("headers", {})
            
            properties: Dict[str, Any] = {}
            for col, header_info in headers.items():
                # Use .get() to safely access dictionary values
                col_name = header_info.get("name", "")
                data_type = header_info.get("data_type", "string")
                
                # Map data types to JSON Schema types
                json_type = self._map_to_json_schema_type(data_type)
                properties[col_name] = {
                    "type": json_type,
                    "description": f"Column {header_info.get('column_letter', '')} - {col_name}"
                }
            
            # Create schema for this sheet
            sheet_schema = {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": properties,
                    "additionalProperties": False
                }
            }
            
            # Add to schema safely
            data_props[sheet_name] = sheet_schema
        
        return schema
    
    def _map_to_json_schema_type(self, data_type: str) -> Union[str, List[str]]:
        """
        Map Excel data types to JSON Schema types.
        
        Args:
            data_type: Excel data type
            
        Returns:
            JSON Schema type or list of types
        """
        # Define the mapping with explicit types
        type_map: Dict[str, Union[str, List[str]]] = {
            "string": "string",
            "integer": "integer",
            "number": "number",
            "boolean": "boolean",
            "datetime": "string",
            "null": ["null", "string"]
        }
        
        # Explicit check if key exists
        if data_type in type_map:
            return type_map[data_type]
        return "string"
    
    def validate_data(self, data: Dict[str, Any]) -> List[str]:
        """
        Validate data against the schema.
        
        Args:
            data: Data to validate
            
        Returns:
            List of validation errors
        """
        errors: List[str] = []
        schema = self.to_json_schema()
        
        try:
            jsonschema.validate(instance={"data": data}, schema=schema)
        except jsonschema.exceptions.ValidationError as e:
            errors.append(str(e))
        
        return errors
    
    def generate_yaml_template(self) -> Dict[str, Any]:
        """
        Generate a YAML template from the schema.
        
        Returns:
            Dictionary with template structure
        """
        template: Dict[str, Any] = {
            "schema": self.structure,
            "data": {}
        }
        
        # Get sheets safely
        sheets = self.structure.get("sheets", {})
        template_data = template["data"]
        
        # Create empty data structure
        for sheet_name, sheet_structure in sheets.items():
            headers = sheet_structure.get("headers", {})
            
            sheet_data: List[Dict[str, Any]] = []
            # Add an example row
            if headers:
                example_row: Dict[str, Any] = {}
                for col, header_info in headers.items():
                    col_name = header_info.get("name", "")
                    # Add placeholder based on data type
                    data_type = header_info.get("data_type", "string")
                    example_row[col_name] = self._get_type_example(data_type)
                
                sheet_data.append(example_row)
            
            # Add to template safely
            template_data[sheet_name] = sheet_data
        
        return template
    
    def _get_type_example(self, data_type: str) -> Any:
        """
        Get an example value for a data type.
        
        Args:
            data_type: Data type
            
        Returns:
            Example value
        """
        examples: Dict[str, Any] = {
            "string": "example",
            "integer": 0,
            "number": 0.0,
            "boolean": False,
            "datetime": "2023-01-01",
            "null": None
        }
        
        # Explicit check if key exists
        if data_type in examples:
            return examples[data_type]
        return "example" 