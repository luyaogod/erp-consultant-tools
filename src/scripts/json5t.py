import json
import json5
from dataclasses import dataclass

@dataclass
class ToolConfig:
    json5_str: str 
    indent: int = 2


def run_tool(config: ToolConfig)->str:
    return json.dumps(json5.loads(config.json5_str), indent=config.indent)


if __name__ == "__main__":
    config = ToolConfig(
        json5_str="""
{
  /** 
   * Smart JSON5 Editor:
   * 1. VS Code-like JSON Editor
   * 2. Error-free editing experience
   * 3. Compatible with JSON/JSON5
   **/ 
  "url": "https://json-5.com",

  // Key without quotes is allowed:
  unquoted: "string",

  // Trailing comma in array is allowed:
  "array1": [1, 2,], 

  // Trailing comma in object is allowed:
  "object1": { "key1": "string", },

  // Missing comma in object is allowed
  "object2": { "key1": "string"  } 

  // Nested data
  "resources":[{"name":"goodeducation","schema":{"fields":[{"name":"content","title":"Content"},{"name":"yescount","title":"YesCount"},{"name":"nocount","title":"NoCount"},{"name":"percentyes","title":"PercentYes"},{"name":"percentno","title":"PercentNo"}]},"format":"csv"},{"name":"pollingdata","schema":{"fields":[{"name":"yes","title":"Yes"},{"name":"no","title":"No"}]},"format":"csv"}]
}
""",
        indent=2,
    )
    
    print(run_tool(config))
    
