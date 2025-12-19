# PCF XmlToJson

A Power Apps Component Framework (PCF) control that lets users upload one or multiple **XML** files and outputs a **JSON string** that can be easily consumed in Canvas Apps using `ParseJSON()`.

<img width="317" height="287" alt="Captura de tela 2025-12-18 204146" src="https://github.com/user-attachments/assets/5b61690f-ab64-4184-9b90-d3b1c5f7d544" />


---

## What this component does

- Upload **1 XML** or **multiple XML files** (configurable).
- Convert XML → JSON and expose the result as an **output property** (string).
- Provide quick actions:
  - **Copy Schema**: copies a JSON Schema generated from the current JSON output.
  - **Copy Schema With Prompt**: copies a ready-to-use ChatGPT prompt **+** the JSON Schema.
- Supports an **external reset** (triggered from the host app) to clear the current state without needing a “Reset” button inside the control.

---
## Usage

Click here to [Download](https://github.com/MarcosMorais95/PCFXmlToJson/releases/tag/v1.0.1) the managed solution.

---
## Key features

- **Single or multi-file mode** via a boolean property.
- **Deterministic output format**: always returns an array (even if only one file is selected).
- **No debug UI** (clean component surface).
- **External reset trigger** (host app controls when to clear).
- **Clipboard utilities** for schema generation and AI assistance.

---

## Output format

The output is a JSON string representing an array of converted files:

```json
[
  {
    "fileName": "Invoice_001.xml",
    "content": { "...converted xml content..." }
  }
]



