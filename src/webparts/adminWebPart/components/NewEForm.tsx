import * as React from 'react';
import { SPFI } from "@pnp/sp";
import {
  TextField,
  PrimaryButton,
  Stack,
  Label,
  Dropdown,
  IDropdownOption
} from "@fluentui/react";
import { IconButton } from "@fluentui/react/lib/Button";

/**
 * Props injected from parent component
 * sp         → Initialized PnP SPFI instance
 * onSuccess  → Callback triggered after successful configuration save
 */
interface NewFormProps {
  sp: SPFI;
  onSuccess: () => void;
}

/**
 * Represents a single field configuration
 * Stored inside ConfigurationJSON in SharePoint
 */
interface ColumnConfig {
  label: string;
  internalName: string;
  type: string;
  required: boolean;
  order: number;
  options?: string[];
  tableName?: string;
  tableRows?: number;
  tableColumns?: number;
  tableHeaders?: string[];
  tableData?: string[][];
  htmlContent?: string;
  tableColumnDefs?: { name: string; type: string; options?: string[] }[];
}

/**
 * Component State
 * Holds both:
 * - Form metadata
 * - Field builder temporary state
 */
interface NewFormState {
  formTitle: string;
  description: string;
  columns: ColumnConfig[];

  newColumnName: string;
  newInternalName: string;
  newColumnType: string;
  required: boolean;
  order: string;

  options: string[];
  newChoiceValue: string;

  // HTML Table specific state
  tableName: string;
  tableRows: string;
  tableColumns: string;
  tableHeaders: string[];
  currentHeaderInput: string;
  tableData: string[][];
  plainHtmlInput: string;
  tableColumnDefs: { name: string; type: string; options?: string[] }[];
  currentHeaderType: string;
  currentHeaderOptions: string[];
  currentHeaderOptionInput: string;

  isFormInfoOpen: boolean;
  isFieldsConfigOpen: boolean;
}

/**
 * NewForm Component
 *
 * Responsibilities:
 * 1. Build dynamic form schema
 * 2. Store configuration in SharePoint list (ConfigList)
 * 3. Provide preview before saving
 */
export default class NewForm extends React.Component<NewFormProps, NewFormState> {

  constructor(props: NewFormProps) {
    super(props);

    /**
     * Initialize state
     */
    this.state = {
      formTitle: "",
      description: "",
      columns: [],
      newColumnName: "",
      newInternalName: "",
      newColumnType: "text",
      required: false,
      order: "",
      options: [],
      newChoiceValue: "",
      tableName: "",
      tableRows: "",
      tableColumns: "",
      tableHeaders: [],
      currentHeaderInput: "",
      tableData: [],
      plainHtmlInput: "",
      tableColumnDefs: [],
      currentHeaderType: "text",
      currentHeaderOptions: [],
      currentHeaderOptionInput: "",
      isFormInfoOpen: true,
      isFieldsConfigOpen: true
    };
  }

  /**
   * Adds a new choice value for dropdown/radio/checkbox
   */
  private _addChoice = () => {
    if (!this.state.newChoiceValue.trim()) return;

    this.setState(prev => ({
      options: [...prev.options, prev.newChoiceValue.trim()],
      newChoiceValue: ""
    }));
  };

  /**
   * Adds a new table header
   */
  private _addTableHeader = (): void => {
    if (!this.state.currentHeaderInput.trim()) return;

    const header = this.state.currentHeaderInput.trim();
    const type = this.state.currentHeaderType || "text";
    const options = this.state.currentHeaderOptions && this.state.currentHeaderOptions.length > 0 ? [...this.state.currentHeaderOptions] : undefined;

    this.setState(prev => ({
      tableHeaders: [...prev.tableHeaders, header],
      tableColumnDefs: [...prev.tableColumnDefs, { name: header, type, options }],
      currentHeaderInput: "",
      currentHeaderType: "text",
      currentHeaderOptions: [],
      currentHeaderOptionInput: ""
    }));
  };

  /**
   * Removes a table header by index
   */
  private _removeTableHeader = (index: number): void => {
    this.setState(prev => ({
      tableHeaders: prev.tableHeaders.filter((_, i) => i !== index),
      tableColumnDefs: prev.tableColumnDefs.filter((_, i) => i !== index)
    }));
  };

  /**
   * Add an option for the current header if it's a choice type
   */
  private _addHeaderOption = (): void => {
    if (!this.state.currentHeaderOptionInput.trim()) return;

    this.setState(prev => ({
      currentHeaderOptions: [...prev.currentHeaderOptions, prev.currentHeaderOptionInput.trim()],
      currentHeaderOptionInput: ""
    }));
  };

  /**
   * Remove an option from current header
   */
  private _removeHeaderOption = (index: number): void => {
    this.setState(prev => ({
      currentHeaderOptions: prev.currentHeaderOptions.filter((_, i) => i !== index)
    }));
  };

  /**
   * Initialize tableData as rows x columns empty strings
   */
  

  /**
   * Set a single cell value in tableData
   */
  private _setTableCell = (r: number, c: number, value: string): void => {
    this.setState(prev => {
      const newData = prev.tableData.map(row => [...row]);
      // Ensure rows
      while (newData.length <= r) {
        newData.push(Array.from({ length: Number(prev.tableColumns) || (c + 1) }, () => ""));
      }
      // Ensure columns in row
      while (newData[r].length <= c) {
        newData[r].push("");
      }
      newData[r][c] = value;
      return { tableData: newData } as Pick<NewFormState, keyof NewFormState>;
    });
  };

  /**
   * Adds a new column configuration to state
   *
   * Handles:
   * - Internal name auto-generation
   * - Order calculation
   * - Attaching options (for choice types)
   * - Table configuration (for htmltable types)
   */
  private _addColumn = (): void => {

    if (!this.state.newColumnName.trim()) return;

    // Validate table-specific requirements
    if (this.state.newColumnType === "htmltable") {
      if (!this.state.tableName.trim()) {
        alert("Table name is required for HTML Table fields.");
        return;
      }
      if (!this.state.tableRows || Number(this.state.tableRows) < 1) {
        alert("Number of rows must be at least 1.");
        return;
      }
      if (!this.state.tableColumns || Number(this.state.tableColumns) < 1) {
        alert("Number of columns must be at least 1.");
        return;
      }
      if (this.state.tableHeaders.length !== Number(this.state.tableColumns)) {
        alert(`Please provide exactly ${this.state.tableColumns} column header(s).`);
        return;
      }
    }

    /**
     * Auto-generate SharePoint-safe internal name
     * Removes spaces and special characters
     */
    const generatedInternalName =
      this.state.newInternalName ||
      this.state.newColumnName
        .replace(/\s+/g, "")
        .replace(/[^a-zA-Z0-9]/g, "");

    /**
     * Determine column order
     * If user provides order → use it
     * Otherwise → auto increment
     */
    let orderNumber: number;

    if (this.state.order) {
      orderNumber = Number(this.state.order);
    } else {
      const maxOrder = this.state.columns.length > 0
        ? Math.max(...this.state.columns.map(c => c.order))
        : 0;
      orderNumber = maxOrder + 1;
    }

    const newColumn: ColumnConfig = {
      label: this.state.newColumnName,
      internalName: generatedInternalName,
      type: this.state.newColumnType,
      required: this.state.required,
      order: orderNumber
    };

    // Attach options for choice-based types
    if (["dropdown", "radio", "checkbox"].includes(this.state.newColumnType)) {
      newColumn.options = [...this.state.options];
    }

    // Attach table configuration for htmltable type
    if (this.state.newColumnType === "htmltable") {
      newColumn.tableName = this.state.tableName;
      newColumn.tableRows = Number(this.state.tableRows);
      newColumn.tableColumns = Number(this.state.tableColumns);
      newColumn.tableHeaders = this.state.tableColumnDefs.map(d => d.name);
      newColumn.tableColumnDefs = this.state.tableColumnDefs ? this.state.tableColumnDefs.map(d => ({ ...d })) : undefined;
      newColumn.tableData = this.state.tableData.map(r => [...r]);
    }
    // Attach table configuration and data for htmlrender type
    if (this.state.newColumnType === "htmlrender") {
      newColumn.tableName = this.state.tableName;
      newColumn.tableRows = Number(this.state.tableRows);
      newColumn.tableColumns = Number(this.state.tableColumns);
      newColumn.tableHeaders = [...this.state.tableHeaders];
      newColumn.tableColumnDefs = this.state.tableColumnDefs ? this.state.tableColumnDefs.map(d => ({ ...d })) : undefined;
      newColumn.tableData = this.state.tableData.map(r => [...r]);
    }
    // Attach html content for plainhtml type
    if (this.state.newColumnType === "plainhtml") {
      newColumn.htmlContent = this.state.plainHtmlInput;
    }

    // Update state immutably
    this.setState(prev => ({
      columns: [...prev.columns, newColumn],
      newColumnName: "",
      newInternalName: "",
      order: "",
      options: [],
      tableName: "",
      tableRows: "",
      tableColumns: "",
      tableHeaders: [],
      currentHeaderInput: "",
      tableData: [],
      plainHtmlInput: ""
    }));
  };

  /**
   * Saves configuration to SharePoint
   *
   * Steps:
   * 1. Validate required metadata
   * 2. Sort fields by order
   * 3. Serialize to JSON
   * 4. Save to ConfigList
   */
  private _submitForm = async (): Promise<void> => {

    if (!this.state.formTitle.trim()) {
      alert("Form Title is required."); 
      return;
    }

    if (this.state.columns.length === 0) {
      alert("Please add at least one field.");
      return;
    }

    try {

      // Always sort before saving to ensure correct rendering order
      const sortedColumns = [...this.state.columns]
        .sort((a, b) => a.order - b.order);

      const finalJson = {
        formTitle: this.state.formTitle,
        description: this.state.description,
        fields: sortedColumns,
      };

      /**
       * Save configuration item
       */
      await this.props.sp.web.lists
        .getByTitle("ConfigList")
        .items.add({
          Title: this.state.formTitle,
          ConfigurationJSON: JSON.stringify(finalJson),
          Status: "Active"
        });

      alert("Form configuration saved successfully!");

      // Reset builder state
      this.setState({
        formTitle: "",
        description: "",
        columns: [],
        newColumnName: "",
        newInternalName: "",
        newColumnType: "text",
        required: false,
        order: "",
        options: [],
        newChoiceValue: "",
        tableName: "",
        tableRows: "",
        tableColumns: "",
        tableHeaders: [],
        currentHeaderInput: ""
      });

      // Notify parent component
      this.props.onSuccess();

    } catch (error) {
      console.error("Error saving configuration:", error);
      alert("Error saving form configuration.");
    }
  };

  /**
   * Render Method
   *
   * Sections:
   * - Form Information
   * - Field Configuration Builder
   * - Live Preview
   * - Submit Button
   */
  public render(): React.ReactElement<NewFormProps> {

    /**
     * Dropdown options
     */
    const columnTypes: IDropdownOption[] = [
      { key: "text", text: "Single line of text" },
      { key: "multiline", text: "Multiple lines of text" },
      { key: "number", text: "Number" },
      { key: "currency", text: "Currency" },
      { key: "date", text: "Date and Time" },
      { key: "dropdown", text: "Dropdown" },
      { key: "radio", text: "Radio Buttons" },
      { key: "checkbox", text: "Checkboxes (Multiple Select)" },
      { key: "yesno", text: "Yes / No" },
      { key: "person", text: "Person or Group" },
      { key: "htmltable", text: "HTML Table" },
      { key: "htmlrender", text: "HTML Render" }
      ,{ key: "plainhtml", text: "Plain HTML" }
    ];

    return (

      <div style={{ margin: "0 auto", padding: "20px" }}>

        <h1>Create New Form</h1>

        {/* ================= FORM DETAILS SECTION ================= */}
        <div style={{
          border: "1px solid #ddd",
          padding: "20px",
          marginBottom: "30px",
          backgroundColor: "#fafafa"
        }}>

          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              cursor: "pointer"
            }}
            onClick={() =>
              this.setState(prev => ({ isFormInfoOpen: !prev.isFormInfoOpen }))
            }
          >
            <h2 style={{ margin: 0 }}>Form Information</h2>

            <IconButton
              iconProps={{
                iconName: this.state.isFormInfoOpen ? "ChevronUp" : "ChevronDown"
              }}
            />
          </div>

          {this.state.isFormInfoOpen && (
            <Stack tokens={{ childrenGap: 15 }} style={{ marginTop: "15px", width: "80%" }}>

              <TextField
                label="Form Title"
                value={this.state.formTitle}
                onChange={(_, v) => this.setState({ formTitle: v || "" })}
                required
              />

              <TextField
                label="Description"
                multiline
                rows={3}
                value={this.state.description}
                onChange={(_, v) => this.setState({ description: v || "" })}
              />

            </Stack>
          )}
        </div>


        {/* ================= COLUMN SECTION ================= */}
        <div style={{
          border: "1px solid #ddd",
          padding: "20px"
        }}>


          <div style={{ marginTop: "20px" }}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                cursor: "pointer"
              }}
              onClick={() =>
                this.setState(prev => ({
                  isFieldsConfigOpen: !prev.isFieldsConfigOpen
                }))
              }
            >
              <h2 style={{ margin: 0 }}>Fields Configuration</h2>

              <IconButton
                iconProps={{
                  iconName: this.state.isFieldsConfigOpen ? "ChevronUp" : "ChevronDown"
                }}
              />
            </div>
            {this.state.isFieldsConfigOpen && (
              <div>
                <h1>Create Column</h1>

                <Stack tokens={{ childrenGap: 15 }} style={{ width: "80%" }}>

                  <TextField
                    label="Column Name"
                    value={this.state.newColumnName}
                    onChange={(_, v) => this.setState({ newColumnName: v || "" })}
                    required
                  />

                  <TextField
                    label="Internal Name"
                    value={this.state.newInternalName}
                    onChange={(_, v) => this.setState({ newInternalName: v || "" })}
                    placeholder="Leave empty to auto-generate"
                  />

                  <Dropdown
                    label="Type of information in this column"
                    selectedKey={this.state.newColumnType}
                    options={columnTypes}
                    onChange={(_, option) =>
                      this.setState({
                        newColumnType: option?.key as string,
                        options: [],
                        newChoiceValue: ""
                      })
                    }
                  />

                  {/* Choice Specific Settings */}
                  {["dropdown", "radio", "checkbox"].includes(this.state.newColumnType) && (
                    <div style={{ padding: "15px", border: "1px solid #eee", borderRadius: "6px" }}>

                      <Label>Add Options</Label>

                      <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <TextField
                          placeholder="Enter option value"
                          value={this.state.newChoiceValue}
                          onChange={(_, v) => this.setState({ newChoiceValue: v || "" })}
                        />
                        <PrimaryButton text="Add" onClick={this._addChoice} />
                      </Stack>

                      <div style={{ marginTop: "10px" }}>
                        {this.state.options.map((c, i) => (
                          <div key={i} style={{
                            background: "#f3f3f3",
                            padding: "5px 10px",
                            borderRadius: "4px",
                            marginBottom: "5px",
                            display: "inline-block",
                            marginRight: "5px"
                          }}>
                            {c}
                          </div>
                        ))}
                      </div>

                    </div>
                  )}

                  {/* Plain HTML Specific Settings */}
                  {this.state.newColumnType === "plainhtml" && (
                    <div style={{ padding: "15px", border: "1px solid #eee", borderRadius: "6px" }}>

                      <Label>Plain HTML Content</Label>

                      <TextField
                        label="HTML"
                        multiline
                        rows={8}
                        placeholder="Paste HTML here (e.g., a <table>...)</"
                        value={this.state.plainHtmlInput}
                        onChange={(_, v) => this.setState({ plainHtmlInput: v || "" })}
                      />

                      <div style={{ marginTop: "10px" }}>
                        <Label>Preview</Label>
                        <div style={{ border: "1px solid #ddd", padding: "10px", marginTop: "6px" }} dangerouslySetInnerHTML={{ __html: this.state.plainHtmlInput || "" }} />
                      </div>

                    </div>
                  )}

                  {/* HTML Render Specific Settings */}
                  {this.state.newColumnType === "htmlrender" && (
                    <div style={{ padding: "15px", border: "1px solid #eee", borderRadius: "6px" }}>

                      <Label>HTML Render Configuration</Label>

                      <div style={{ textAlign: "center", marginBottom: "10px" }}>
                        <TextField
                          label="Table Name"
                          placeholder="e.g., Employee Details"
                          value={this.state.tableName}
                          onChange={(_, v) => this.setState({ tableName: v || "" })}
                        />
                      </div>

                      <Stack horizontal tokens={{ childrenGap: 15 }} style={{ marginBottom: "15px" }}>
                        <TextField
                          label="Number of Rows (data rows)"
                          type="number"
                          min="1"
                          max="200"
                          value={this.state.tableRows}
                          onChange={(_, v) => this.setState({ tableRows: v || "" })}
                          style={{ width: "180px" }}
                        />
                        <TextField
                          label="Number of Columns"
                          type="number"
                          min="1"
                          max="50"
                          value={this.state.tableColumns}
                          onChange={(_, v) => {
                            const cols = Number(v || 0);
                            this.setState(prev => ({
                              tableColumns: v || "",
                              tableHeaders: [],
                              currentHeaderInput: "",
                              currentHeaderType: "",
                              currentHeaderOptions: [],
                              currentHeaderOptionInput: "",
                              tableColumnDefs: [],
                              tableData: cols > 0 && prev.tableRows ? Array.from({ length: Number(prev.tableRows) }, () => Array.from({ length: cols }, () => "")) : []
                            }));
                          }}
                          style={{ width: "180px" }}
                        />
                      </Stack>

                      {/* Column Headers Input */}
                      {this.state.tableColumns && Number(this.state.tableColumns) > 0 && (
                        <div style={{ marginTop: "10px" }}>
                          <Label>Column Headers ({this.state.tableHeaders.length} of {this.state.tableColumns})</Label>

                          <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginBottom: "10px" }}>
                                <TextField
                                  placeholder="Enter column header"
                                  value={this.state.currentHeaderInput}
                                  onChange={(_, v) => this.setState({ currentHeaderInput: v || "" })}
                                />
                                <Dropdown
                                  selectedKey={this.state.currentHeaderType}
                                  options={[
                                    { key: 'text', text: 'Single line of text' },
                                    { key: 'multiline', text: 'Multiple lines of text' },
                                    { key: 'number', text: 'Number' },
                                    { key: 'currency', text: 'Currency' },
                                    { key: 'date', text: 'Date' },
                                    { key: 'dropdown', text: 'Dropdown' },
                                    { key: 'radio', text: 'Radio' },
                                    { key: 'checkbox', text: 'Checkbox' },
                                    { key: 'yesno', text: 'Yes / No' }
                                  ]}
                                  onChange={(_, option) => {
                                    this.setState({
                                      currentHeaderType: option?.key as string,
                                      currentHeaderOptions: [],
                                      currentHeaderOptionInput: ""
                                    });
                                  }}
                                  styles={{ root: { width: '200px' } }}
                                />

                                {/* Options input for choice types */}
                                {['dropdown', 'radio', 'checkbox'].includes(this.state.currentHeaderType) && (
                                  <div style={{ marginTop: 12, padding: 12, backgroundColor: '#f9f9f9', borderRadius: 4 }}>
                                    <Label>Add options for this column</Label>
                                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                                      <TextField
                                        placeholder="Enter option"
                                        value={this.state.currentHeaderOptionInput}
                                        onChange={(_, v) => this.setState({ currentHeaderOptionInput: v || "" })}
                                      />
                                      <PrimaryButton
                                        text="Add Option"
                                        onClick={this._addHeaderOption}
                                      />
                                    </Stack>
                                    <div style={{ marginTop: 8 }}>
                                      {this.state.currentHeaderOptions.map((opt, idx) => (
                                        <div key={idx} style={{
                                          background: '#e8f4f8',
                                          padding: '4px 8px',
                                          borderRadius: 4,
                                          marginBottom: 4,
                                          display: 'inline-block',
                                          marginRight: 4
                                        }}>
                                          {opt}
                                          <IconButton
                                            iconProps={{ iconName: 'Cancel' }}
                                            onClick={() => this._removeHeaderOption(idx)}
                                            styles={{ root: { height: 20, width: 20, marginLeft: 4 } }}
                                          />
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                )}

                                <PrimaryButton
                                  text="Add Header"
                                  onClick={this._addTableHeader}
                                  disabled={
                                    this.state.tableColumnDefs.length >= Number(this.state.tableColumns) ||
                                    (['dropdown', 'radio', 'checkbox'].includes(this.state.currentHeaderType) && this.state.currentHeaderOptions.length === 0)
                                  }
                                />
                          </Stack>

                          <div style={{ marginTop: "10px" }}>
                            {this.state.tableColumnDefs.map((def, i) => (
                              <div key={i} style={{
                                background: "#e8f4f8",
                                padding: "8px 12px",
                                borderRadius: "4px",
                                marginBottom: "5px",
                                display: "flex",
                                justifyContent: "space-between",
                                alignItems: "center"
                              }}>
                                <div>
                                  <div style={{ fontWeight: 600 }}>{i + 1}. {def.name}</div>
                                  <div style={{ fontSize: 12, color: '#444' }}>{def.type}{def.options ? ` — ${def.options.join(', ')}` : ''}</div>
                                </div>
                                <IconButton
                                  iconProps={{ iconName: "Delete" }}
                                  onClick={() => this._removeTableHeader(i)}
                                  styles={{ root: { height: "24px", width: "24px" } }}
                                />
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      {/* Cell Data Grid */}
                      {this.state.tableColumnDefs.length > 0 && this.state.tableColumns && this.state.tableRows && Number(this.state.tableColumns) > 0 && Number(this.state.tableRows) > 0 && (
                        <div style={{ marginTop: "15px" }}>
                          <Label>Enter cell data</Label>
                          <div style={{ overflowX: "auto", marginTop: "8px" }}>
                            {Array.from({ length: Number(this.state.tableRows) }).map((_, r) => (
                              <div key={r} style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
                                {Array.from({ length: Number(this.state.tableColumns) }).map((_, c) => {
                                  const colDef = this.state.tableColumnDefs[c] || { name: `Col ${c + 1}`, type: 'text' };
                                  let val = (this.state.tableData[r] && this.state.tableData[r][c]) || "";
                                  
                                  // For yesno type, default to 'false' if empty
                                  if (colDef.type === 'yesno' && !val) {
                                    val = 'false';
                                  }
                                  
                                  // Render appropriate input per column type
                                  if (colDef.type === 'number' || colDef.type === 'currency') {
                                    return (
                                      <input key={`${r}-${c}`} type="number" value={val} onChange={(e) => this._setTableCell(r, c, e.target.value)} style={{ width: 160, padding: '6px' }} />
                                    );
                                  }
                                  if (colDef.type === 'date') {
                                    return (
                                      <input key={`${r}-${c}`} type="date" value={val} onChange={(e) => this._setTableCell(r, c, e.target.value)} style={{ width: 160, padding: '6px' }} />
                                    );
                                  }
                                  if (colDef.type === 'dropdown') {
                                    return (
                                      <select key={`${r}-${c}`} value={val} onChange={(e) => this._setTableCell(r, c, e.target.value)} style={{ width: 160, padding: '6px' }}>
                                        <option value="">Select</option>
                                        {(colDef.options || []).map((o, idx) => <option key={idx} value={o}>{o}</option>)}
                                      </select>
                                    );
                                  }
                                  if (colDef.type === 'radio') {
                                    return (
                                      <div key={`${r}-${c}`} style={{ display: 'flex', gap: 6 }}>
                                        {(colDef.options || []).map((o, idx) => (
                                          <label key={idx}><input type="radio" name={`cell-${r}-${c}`} value={o} checked={val===o} onChange={(e)=>this._setTableCell(r,c,e.target.value)} />{o}</label>
                                        ))}
                                      </div>
                                    );
                                  }
                                  if (colDef.type === 'checkbox') {
                                    // store comma-separated selections
                                    const selections = val ? val.split(',') : [];
                                    return (
                                      <div key={`${r}-${c}`} style={{ display: 'flex', gap: 6, flexDirection: 'column' }}>
                                        {(colDef.options || []).map((o, idx) => (
                                          <label key={idx}><input type="checkbox" checked={selections.includes(o)} onChange={(e)=>{
                                            const newSel = e.target.checked ? [...selections, o] : selections.filter(s=>s!==o);
                                            this._setTableCell(r,c,newSel.join(','));
                                          }} />{o}</label>
                                        ))}
                                      </div>
                                    );
                                  }
                                  if (colDef.type === 'yesno') {
                                    return (
                                      <input key={`${r}-${c}`} type="checkbox" checked={val==='true'} onChange={(e)=>this._setTableCell(r,c,e.target.checked? 'true':'false')} />
                                    );
                                  }
                                  // default to text/multiline
                                  if (colDef.type === 'multiline') {
                                    return (
                                      <textarea key={`${r}-${c}`} value={val} onChange={(e)=>this._setTableCell(r,c,e.target.value)} style={{ width: 300, padding: 6 }} />
                                    );
                                  }
                                  return (
                                    <input key={`${r}-${c}`} type="text" value={val} onChange={(e) => this._setTableCell(r, c, e.target.value)} style={{ width: 160, padding: '6px' }} />
                                  );
                                })}
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                    </div>
                  )}

                  {/* HTML Table Specific Settings */}
                  {this.state.newColumnType === "htmltable" && (
                    <div style={{ padding: "15px", border: "1px solid #eee", borderRadius: "6px" }}>

                      <Label>Table Configuration</Label>

                      <div style={{ textAlign: "center", marginBottom: "10px" }}>
                        <TextField
                          label="Table Name"
                          placeholder="e.g., Employee Details"
                          value={this.state.tableName}
                          onChange={(_, v) => this.setState({ tableName: v || "" })}
                        />
                      </div>

                      <Stack horizontal tokens={{ childrenGap: 15 }} style={{ marginBottom: "15px" }}>
                        <TextField
                          label="Number of Rows"
                          type="number"
                          min="1"
                          max="50"
                          value={this.state.tableRows}
                          onChange={(_, v) => this.setState({ tableRows: v || "" })}
                          style={{ width: "150px" }}
                        />
                        <TextField
                          label="Number of Columns"
                          type="number"
                          min="1"
                          max="50"
                          value={this.state.tableColumns}
                          onChange={(_, v) => {
                            this.setState({
                              tableColumns: v || "",
                              tableHeaders: [],
                              currentHeaderInput: "",
                              tableColumnDefs: [],
                              currentHeaderType: "",
                              currentHeaderOptions: [],
                              currentHeaderOptionInput: ""
                            });
                          }}
                          style={{ width: "150px" }}
                        />
                      </Stack>

                      {/* Column Headers Input */}
                      {this.state.tableColumns && Number(this.state.tableColumns) > 0 && (
                        <div style={{ marginTop: "10px" }}>
                          <Label>Column Headers ({this.state.tableColumnDefs.length} of {this.state.tableColumns})</Label>

                          <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginBottom: "10px" }}>
                                <TextField
                                  placeholder="Enter column header"
                                  value={this.state.currentHeaderInput}
                                  onChange={(_, v) => this.setState({ currentHeaderInput: v || "" })}
                                />
                                <Dropdown
                                  selectedKey={this.state.currentHeaderType}
                                  options={[
                                    { key: 'text', text: 'Single line of text' },
                                    { key: 'multiline', text: 'Multiple lines of text' },
                                    { key: 'number', text: 'Number' },
                                    { key: 'currency', text: 'Currency' },
                                    { key: 'date', text: 'Date' },
                                    { key: 'dropdown', text: 'Dropdown' },
                                    { key: 'radio', text: 'Radio' },
                                    { key: 'checkbox', text: 'Checkbox' },
                                    { key: 'yesno', text: 'Yes / No' }
                                  ]}
                                  onChange={(_, option) => {
                                    this.setState({
                                      currentHeaderType: option?.key as string,
                                      currentHeaderOptions: [],
                                      currentHeaderOptionInput: ""
                                    });
                                  }}
                                  styles={{ root: { width: '200px' } }}
                                />

                                {/* Options input for choice types */}
                                {['dropdown', 'radio', 'checkbox'].includes(this.state.currentHeaderType) && (
                                  <div style={{ marginTop: 12, padding: 12, backgroundColor: '#f9f9f9', borderRadius: 4 }}>
                                    <Label>Add options for this column</Label>
                                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                                      <TextField
                                        placeholder="Enter option"
                                        value={this.state.currentHeaderOptionInput}
                                        onChange={(_, v) => this.setState({ currentHeaderOptionInput: v || "" })}
                                      />
                                      <PrimaryButton
                                        text="Add Option"
                                        onClick={this._addHeaderOption}
                                      />
                                    </Stack>
                                    <div style={{ marginTop: 8 }}>
                                      {this.state.currentHeaderOptions.map((opt, idx) => (
                                        <div key={idx} style={{
                                          background: '#e8f4f8',
                                          padding: '4px 8px',
                                          borderRadius: 4,
                                          marginBottom: 4,
                                          display: 'inline-block',
                                          marginRight: 4
                                        }}>
                                          {opt}
                                          <IconButton
                                            iconProps={{ iconName: 'Cancel' }}
                                            onClick={() => this._removeHeaderOption(idx)}
                                            styles={{ root: { height: 20, width: 20, marginLeft: 4 } }}
                                          />
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                )}

                                <PrimaryButton
                                  text="Add Header"
                                  onClick={this._addTableHeader}
                                  disabled={
                                    this.state.tableColumnDefs.length >= Number(this.state.tableColumns) ||
                                    (['dropdown', 'radio', 'checkbox'].includes(this.state.currentHeaderType) && this.state.currentHeaderOptions.length === 0)
                                  }
                                />
                          </Stack>

                          <div style={{ marginTop: "10px" }}>
                            {this.state.tableColumnDefs.map((def, i) => (
                              <div key={i} style={{
                                background: "#e8f4f8",
                                padding: "8px 12px",
                                borderRadius: "4px",
                                marginBottom: "5px",
                                display: "flex",
                                justifyContent: "space-between",
                                alignItems: "center"
                              }}>
                                <div>
                                  <div style={{ fontWeight: 600 }}>{i + 1}. {def.name}</div>
                                  <div style={{ fontSize: 12, color: '#444' }}>{def.type}{def.options ? ` — ${def.options.join(', ')}` : ''}</div>
                                </div>
                                <IconButton
                                  iconProps={{ iconName: "Delete" }}
                                  onClick={() => this._removeTableHeader(i)}
                                  styles={{ root: { height: "24px", width: "24px" } }}
                                />
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                    </div>
                  )}

                  <TextField
                    label="Column Serial Order (Position)"
                    type="number"
                    value={this.state.order}
                    onChange={(_, v) => this.setState({ order: v || "" })}
                    placeholder="Leave empty for auto ordering"
                  />


                  {/* Required */}
                  <div>
                    <Label>Require that this column contains information?</Label>
                    <input
                      type="checkbox"
                      checked={this.state.required}
                      onChange={(e) =>
                        this.setState({ required: e.target.checked })
                      }
                    /> Required
                  </div>

                  <PrimaryButton text="Add Column" onClick={this._addColumn}
                    disabled={
                      !this.state.newColumnName ||
                      (
                        ["dropdown", "radio", "checkbox"].includes(this.state.newColumnType) &&
                        this.state.options.length === 0
                      ) ||
                      (
                        this.state.newColumnType === "htmltable" &&
                        (
                          !this.state.tableName ||
                          !this.state.tableRows ||
                          !this.state.tableColumns ||
                          this.state.tableHeaders.length !== Number(this.state.tableColumns)
                        )
                      )
                    }
                  />

                </Stack>

                {/* Preview */}
                {[...this.state.columns]
                  .sort((a, b) => a.order - b.order)
                  .map((col, i) => (
                    <div key={i} style={{ padding: "8px", borderBottom: "1px solid #eee" }}>
                      <strong>{col.order}. {col.label}</strong> —
                      {col.type === "dropdown" && "Dropdown"}
                      {col.type === "radio" && "Radio Buttons"}
                      {col.type === "checkbox" && "Checkboxes"}
                      {col.type === "text" && "Single line of text"}
                      {col.type === "multiline" && "Multiple lines of text"}
                      {col.type === "number" && "Number"}
                      {col.type === "currency" && "Currency"}
                      {col.type === "date" && "Date and Time"}
                      {col.type === "yesno" && "Yes / No"}
                      {col.type === "person" && "Person or Group"}
                      {col.type === "htmltable" && "HTML Table"}
                        {col.type === "htmlrender" && "HTML Render"}
                        {col.type === "plainhtml" && "Plain HTML"}
                      <div style={{ fontSize: "12px", color: "#666" }}>
                        Internal Name: {col.internalName}
                      </div>
                      {col.required && " (Required)"}
                      {col.options && (
                        <div>
                          Choices: {col.options.join(", ")}
                        </div>
                      )}
                      {col.tableName && (
                        <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
                          Table: {col.tableName} ({col.tableRows} rows × {col.tableColumns} columns)<br />
                          Headers: {col.tableHeaders?.join(", ")}
                        </div>
                      )}
                    </div>
                  ))}

                {/* Submit Button */}
                <div style={{ textAlign: "center", marginTop: "40px", width: "80%" }}>
                  <PrimaryButton
                    text="Submit Form Configuration"
                    onClick={this._submitForm}
                    styles={{ root: { padding: "0 30px", height: "40px" } }}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }
}
