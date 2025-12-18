import { IInputs, IOutputs } from "./generated/ManifestTypes";

// Recursive type that represents the JSON generated from XML
type JsonPrimitive = string | number | boolean | null;
type JsonValue = JsonPrimitive | JsonObject | JsonValue[];
interface JsonObject {
  [key: string]: JsonValue;
}

// Structure that represents each converted file
interface ConvertedFile {
  fileName: string;
  content?: JsonValue;
  error?: string;
}

const POWER_APPS_PROMPT =
  "Create a Power Apps collection called colName by parsing the JSON below using ParseJSON.\n\n" +
  "Extract relevant fields using AddColumns, convert values using Text() or Value(), and use DropColumns to remove the original Value column (do not use quotes around Value).\n\n" +
  "Replace Self.jsonResult with the actual source if needed (e.g., ThisItem.jsonResult).\n\n" +
  "Here's the JSON:";

export class XmlToJson implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

  private root!: HTMLDivElement;
  private fileInput!: HTMLInputElement;

  private resultContainer!: HTMLDivElement;
  private jsonTextArea!: HTMLTextAreaElement;

  private copyButton!: HTMLButtonElement;
  private copyWithPromptButton!: HTMLButtonElement;

  // Output-only value (do NOT read from context.parameters)
  private jsonResult?: string;

  private resetValue?: number;
  private allowMultiple = true;
  private isSchemaVisible = false;

  // Stable handlers for removeEventListener
  private onFileInputChangeHandler = (): void => this.onFileInputChange();
  private onCopyClickHandler = (): void => this.copySchema();
  private onCopyWithPromptClickHandler = (): void => this.copySchemaWithPrompt();

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;

    // Read initial INPUT parameters
    this.resetValue = context.parameters.Reset.raw ?? undefined;
    this.allowMultiple = context.parameters.AllowMultiple.raw ?? true;
    this.isSchemaVisible = context.parameters.IsSchemaVisible.raw ?? false;

    // Root wrapper
    this.root = document.createElement("div");
    this.root.style.display = "flex";
    this.root.style.flexDirection = "column";
    this.root.style.width = "100%";

    // File input
    this.fileInput = document.createElement("input");
    this.fileInput.type = "file";
    this.fileInput.accept = ".xml,text/xml,application/xml";
    this.fileInput.multiple = this.allowMultiple;
    this.fileInput.style.width = "100%";
    this.fileInput.addEventListener("change", this.onFileInputChangeHandler);

    // Result UI
    this.resultContainer = document.createElement("div");
    this.resultContainer.style.display = "flex";
    this.resultContainer.style.flexDirection = "column";
    this.resultContainer.style.width = "100%";
    this.resultContainer.style.marginTop = "8px";

    // Button row
    const buttonRow = document.createElement("div");
    buttonRow.style.display = "flex";
    buttonRow.style.flexDirection = "row";
    buttonRow.style.gap = "8px";
    buttonRow.style.alignItems = "center";
    buttonRow.style.flexWrap = "wrap";

    this.copyButton = document.createElement("button");
    this.copyButton.type = "button";
    this.copyButton.textContent = "Copy Schema";
    this.copyButton.addEventListener("click", this.onCopyClickHandler);

    this.copyWithPromptButton = document.createElement("button");
    this.copyWithPromptButton.type = "button";
    this.copyWithPromptButton.textContent = "Copy Schema With Prompt";
    this.copyWithPromptButton.addEventListener("click", this.onCopyWithPromptClickHandler);

    buttonRow.appendChild(this.copyButton);
    buttonRow.appendChild(this.copyWithPromptButton);

    this.jsonTextArea = document.createElement("textarea");
    this.jsonTextArea.readOnly = true;
    this.jsonTextArea.style.width = "100%";
    this.jsonTextArea.style.minHeight = "160px";
    this.jsonTextArea.style.marginTop = "4px";
    this.jsonTextArea.value = this.jsonResult ?? "";

    this.resultContainer.appendChild(buttonRow);
    this.resultContainer.appendChild(this.jsonTextArea);

    this.root.appendChild(this.fileInput);
    this.root.appendChild(this.resultContainer);

    this.container.appendChild(this.root);

    this.updateJsonDisplay();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Detect Reset changes (external trigger)
    const newReset = context.parameters.Reset.raw ?? undefined;
    if (newReset !== this.resetValue) {
      this.resetValue = newReset;
      if (newReset !== null && newReset !== undefined) {
        this.resetControl();
      }
    }

    // Detect AllowMultiple changes
    const newAllowMultiple = context.parameters.AllowMultiple.raw ?? true;
    if (newAllowMultiple !== this.allowMultiple) {
      this.allowMultiple = newAllowMultiple;
      if (this.fileInput) {
        this.fileInput.multiple = this.allowMultiple;
      }
    }

    // Detect visibility changes for the JSON visualization
    const newSchemaVisibility = context.parameters.IsSchemaVisible.raw ?? false;
    if (newSchemaVisibility !== this.isSchemaVisible) {
      this.isSchemaVisible = newSchemaVisibility;
      this.updateJsonDisplay();
    }
  }

  public getOutputs(): IOutputs {
    return {
      jsonResult: this.jsonResult
    };
  }

  public destroy(): void {
    if (this.fileInput) {
      this.fileInput.removeEventListener("change", this.onFileInputChangeHandler);
    }
    if (this.copyButton) {
      this.copyButton.removeEventListener("click", this.onCopyClickHandler);
    }
    if (this.copyWithPromptButton) {
      this.copyWithPromptButton.removeEventListener("click", this.onCopyWithPromptClickHandler);
    }
  }

  // ============================================================
  // Events
  // ============================================================

  private onFileInputChange(): void {
    const files = this.fileInput.files;

    if (!files || files.length === 0) {
      this.jsonResult = undefined;
      this.updateJsonDisplay();
      this.notifyOutputChanged();
      return;
    }

    if (!this.allowMultiple) {
      const first = files[0];
      if (first) void this.processFiles([first]);
      return;
    }

    void this.processFiles(files);
  }

  // ============================================================
  // Core logic
  // ============================================================

  private resetControl(): void {
    if (this.fileInput) {
      this.fileInput.value = "";
    }
    this.jsonResult = undefined;
    this.updateJsonDisplay();
    this.notifyOutputChanged();
  }

  private async processFiles(files: FileList | File[]): Promise<void> {
    const fileArray: File[] = Array.isArray(files) ? files : Array.from(files);

    try {
      const results = await Promise.all(fileArray.map((f) => this.readAndConvertFile(f)));
      this.jsonResult = JSON.stringify(results, null, 2);
    } catch (e) {
      const message = e instanceof Error && e.message ? e.message : String(e);
      this.jsonResult = JSON.stringify(
        { error: "Error processing files", detail: message },
        null,
        2
      );
    }

    this.updateJsonDisplay();
    this.notifyOutputChanged();
  }

  private readAndConvertFile(file: File): Promise<ConvertedFile> {
    return new Promise((resolve) => {
      const reader = new FileReader();

      reader.onload = () => {
        try {
          const xmlText = String(reader.result ?? "");
          const obj = this.convertXmlToObject(xmlText);
          resolve({ fileName: file.name, content: obj });
        } catch (e) {
          const message = e instanceof Error && e.message ? e.message : "Error parsing XML";
          resolve({ fileName: file.name, error: message });
        }
      };

      reader.onerror = () => resolve({ fileName: file.name, error: "Error reading file" });

      reader.readAsText(file);
    });
  }

  private convertXmlToObject(xmlString: string): JsonValue {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "application/xml");

    const parserError = doc.getElementsByTagName("parsererror");
    if (parserError && parserError.length > 0) {
      throw new Error("Invalid XML");
    }

    return this.nodeToObject(doc.documentElement);
  }

  private nodeToObject(node: Element): JsonValue {
    const obj: JsonObject = {};

    if (node.attributes && node.attributes.length > 0) {
      const attrs: JsonObject = {};
      for (const attr of Array.from(node.attributes)) {
        attrs[attr.name] = attr.value;
      }
      obj["@attributes"] = attrs;
    }

    const childNodes = node.childNodes;
    let hasElementChildren = false;
    let textContent = "";

    for (const child of Array.from(childNodes)) {
      if (child.nodeType === Node.ELEMENT_NODE) {
        hasElementChildren = true;

        const childElement = child as Element;
        const childObj = this.nodeToObject(childElement);
        const name = childElement.nodeName;

        const existing = obj[name];
        if (existing === undefined) obj[name] = childObj;
        else if (Array.isArray(existing)) existing.push(childObj);
        else obj[name] = [existing, childObj];
      } else if (child.nodeType === Node.TEXT_NODE) {
        const value = child.textContent?.trim();
        if (value) textContent += value;
      }
    }

    if (!hasElementChildren) {
      const hasOnlyAttributes =
        Object.keys(obj).length === 0 ||
        (Object.keys(obj).length === 1 &&
          Object.prototype.hasOwnProperty.call(obj, "@attributes"));

      if (textContent && hasOnlyAttributes) {
        if (Object.prototype.hasOwnProperty.call(obj, "@attributes")) {
          obj["#text"] = textContent;
          return obj;
        }
        return textContent;
      }

      if (textContent) obj["#text"] = textContent;
      return obj;
    }

    if (textContent) obj["#text"] = textContent;
    return obj;
  }

  // ============================================================
  // Copy helpers
  // ============================================================

  private copySchema(): void {
    const textToCopy = this.jsonResult ?? "";
    if (!textToCopy) return;
    this.writeToClipboard(textToCopy);
  }

  private copySchemaWithPrompt(): void {
    const schema = this.jsonResult ?? "";
    if (!schema) return;

    const textToCopy = `${POWER_APPS_PROMPT}\n\n${schema}`;
    this.writeToClipboard(textToCopy);
  }

  private writeToClipboard(text: string): void {
    // Prefer modern clipboard API when available
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => {
        this.fallbackCopy(text);
      });
      return;
    }

    this.fallbackCopy(text);
  }

  private fallbackCopy(text: string): void {
    const temp = document.createElement("textarea");
    temp.value = text;
    temp.setAttribute("readonly", "true");
    temp.style.position = "fixed";
    temp.style.left = "-9999px";
    temp.style.top = "0";

    document.body.appendChild(temp);
    temp.select();
    try {
      document.execCommand("copy");
    } finally {
      document.body.removeChild(temp);
    }
  }

  // ============================================================
  // UI helpers
  // ============================================================

  private updateJsonDisplay(): void {
    if (this.jsonTextArea) {
      this.jsonTextArea.value = this.jsonResult ?? "";
    }

    const shouldShow = this.isSchemaVisible;
    const hasValue = !!this.jsonResult;

    if (this.resultContainer) {
      this.resultContainer.style.display = shouldShow ? "flex" : "none";
    }

    if (this.copyButton) {
      this.copyButton.style.display = shouldShow ? "inline-block" : "none";
      this.copyButton.disabled = !hasValue;
    }

    if (this.copyWithPromptButton) {
      this.copyWithPromptButton.style.display = shouldShow ? "inline-block" : "none";
      this.copyWithPromptButton.disabled = !hasValue;
    }
  }
}
