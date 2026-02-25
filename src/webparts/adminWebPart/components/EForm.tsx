import * as React from 'react';
import { SPFI } from "@pnp/sp";

/**
 * Props for Eform component
 * sp - Initialized PnP SPFI instance (injected from parent / SPFx context)
 * formIDValue - ID of the configuration item stored in SharePoint list
 */
interface AdminFormProps {
    sp: SPFI,
    formIDValue: string
}

/**
 * Component state
 * formData - Stores user input values
 * config - Stores configuration item retrieved from SharePoint
 * tableConfig - Stores table configuration (rows, columns) for each table field
 */
interface AdminFormState {
    formData: any;
    config: any;
    tableConfig: { [key: string]: { rows: number; columns: number } };
}

/**
 * Eform
 * Dynamically renders form fields based on ConfigurationJSON stored in SharePoint
 */
export default class Eform extends React.Component<AdminFormProps, AdminFormState> {

    constructor(props: AdminFormProps) {
        super(props);

        // Initialize component state
        this.state = {
            formData: {},
            config: null,
            tableConfig: {},
        };
    }

    /**
     * Lifecycle method
     * Loads configuration data when component mounts
     */
    public async componentDidMount(): Promise<void> {
        await this._loadItems();
    }

    /**
     * Fetch configuration item from SharePoint list
     */
    private async _loadItems(): Promise<void> {
        try {

            // Validate ID before calling SharePoint
            if (!this.props.formIDValue || isNaN(Number(this.props.formIDValue))) {
                console.error("Invalid form ID provided.");
                return;
            }

            const item = await this.props.sp.web.lists
                .getByTitle("ConfigList")
                .items
                .getById(Number(this.props.formIDValue))();

            this.setState({ config: item });

        } catch (error) {

            console.error("Error fetching configuration item:", error);
        }
    }

    /**
     * Dynamically renders field based on type
     * Field schema comes from ConfigurationJSON
     */
    private renderDynamicField = (field: any) => {

        // Shared styling 
        const fieldContainerStyle: React.CSSProperties = {
            marginBottom: "20px",
            display: "flex",
            flexDirection: "column"
        };

        const labelStyle: React.CSSProperties = {
            marginBottom: "6px",
            fontWeight: 600
        };

        const inputStyle: React.CSSProperties = {
            padding: "8px 10px",
            borderRadius: "4px",
            border: "1px solid #ccc",
            fontSize: "14px",
            width: "80%"
        };

        switch (field.type) {

            /* ===========================
               BASIC FIELDS
            ============================ */
            //normal label to be added

            case "text":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>
                            {field.label}
                            {field.required && <span style={{ color: "red" }}> *</span>}
                        </label>
                        <input
                            type="text"
                            name={field.internalName}
                            required={field.required}
                            style={inputStyle}
                        />
                    </div>
                );

            case "number":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <input
                            type="number"
                            name={field.internalName}
                            style={inputStyle}
                        />
                    </div>
                );

            case "date":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <input
                            type="date"
                            name={field.internalName}
                            style={inputStyle}
                        />
                    </div>
                );

            case "multiline": //Multiline to be changed to Textarea
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <textarea
                            name={field.internalName}
                            rows={4}
                            style={{ ...inputStyle, resize: "vertical" }}
                        />
                    </div>
                );

            /* ===========================
               CHOICE TYPES
            ============================ */

            case "dropdown":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <select name={field.internalName} style={inputStyle}>
                            <option value="">Select</option>
                            {field.options?.map((opt: string, i: number) => (
                                <option key={`${field.internalName}-${i}`} value={opt}>
                                    {opt}
                                </option>
                            ))}
                        </select>
                    </div>
                );

            case "radio":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        {field.options?.map((opt: string, i: number) => (
                            <label key={`${field.internalName}-${i}`} style={{ marginBottom: "4px" }}>
                                <input
                                    type="radio"
                                    name={field.internalName}
                                    value={opt}
                                    style={{ marginRight: "6px" }}
                                />
                                {opt}
                            </label>
                        ))}
                    </div>
                );

            case "checkbox":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        {field.options?.map((opt: string, i: number) => (
                            <label key={`${field.internalName}-${i}`} style={{ marginBottom: "4px" }}>
                                <input
                                    type="checkbox"
                                    name={field.internalName}
                                    value={opt}
                                    style={{ marginRight: "6px" }}
                                />
                                {opt}
                            </label>
                        ))}
                    </div>
                );

            case "yesno":
                return (
                    <div key={field.internalName} style={{ marginBottom: "20px" }}>
                        <label>
                            <input
                                type="checkbox"
                                name={field.internalName}
                                style={{ marginRight: "6px" }}
                            />
                            {field.label}
                        </label>
                    </div>
                );

            /* ===========================
               SHAREPOINT-LIKE TYPES
            ============================ */

            case "currency":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <input
                            type="number"
                            step="0.01"
                            name={field.internalName}
                            style={inputStyle}
                        />
                    </div>
                );

            case "htmltable":
                // Use per-column definitions (name + type + options) to render typed inputs bound to formData
                const rows = Number(field.tableRows || 0);
                const cols = Number(field.tableColumns || 0);
                const colDefs = field.tableColumnDefs || (field.tableHeaders || []).map((h: string) => ({ name: h, type: 'text' }));

                const renderCell = (r: number, c: number) => {
                    const name = `${field.internalName}-${r}-${c}`;
                    const def = colDefs[c] || { name: `Column ${c + 1}`, type: 'text' };
                    let val = this.state.formData?.[name] ?? (field.tableData && field.tableData[r] ? field.tableData[r][c] : "");

                    // For yesno type, default to 'false' if empty
                    if (def.type === 'yesno' && !val) {
                        val = 'false';
                    }

                    const setVal = (v: any) => {
                        this.setState(prev => ({ formData: { ...(prev.formData || {}), [name]: v } }));
                    };

                    if (def.type === 'number' || def.type === 'currency') {
                        return <input type="number" value={val || ''} onChange={e => setVal(e.target.value)} style={{ width: '100%', padding: 6 }} name={name} />;
                    }
                    if (def.type === 'date') {
                        return <input type="date" value={val || ''} onChange={e => setVal(e.target.value)} style={{ width: '100%', padding: 6 }} name={name} />;
                    }
                    if (def.type === 'dropdown' || def.type === 'radio') {
                        return (
                            <select value={val || ''} onChange={e => setVal(e.target.value)} style={{ width: '100%', padding: 6 }} name={name}>
                                <option value="">Select</option>
                                {(def.options || []).map((o: string, i: number) => <option key={i} value={o}>{o}</option>)}
                            </select>
                        );
                    }
                    if (def.type === 'checkbox') {
                        const selections = val ? String(val).split(',') : [];
                        return (
                            <div style={{ display: 'flex', gap: 8, flexDirection: 'column' }}>
                                {(def.options || []).map((o: string, i: number) => (
                                    <label key={i}><input type="checkbox" checked={selections.includes(o)} onChange={(e) => {
                                        const newSel = e.target.checked ? [...selections, o] : selections.filter(s => s !== o);
                                        setVal(newSel.join(','));
                                    }} /> {o}</label>
                                ))}
                            </div>
                        );
                    }
                    if (def.type === 'yesno') {
                        return <input type="checkbox" checked={val === 'true' || val === true} onChange={e => setVal(e.target.checked ? 'true' : 'false')} name={name} />;
                    }
                    if (def.type === 'multiline') {
                        return <textarea value={val || ''} onChange={e => setVal(e.target.value)} style={{ width: '100%', padding: 6 }} name={name} />;
                    }

                    // default text
                    return <input type="text" value={val || ''} onChange={e => setVal(e.target.value)} style={{ width: '100%', padding: 6 }} name={name} />;
                };

                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <div style={{ border: '1px solid #000000', padding: '10px', borderRadius: '4px' }}>
                            <table style={{ borderCollapse: 'collapse', width: '100%', marginTop: '10px', border: '1px solid #ccc' }}>
                                <thead>
                                    <tr style={{ textAlign: 'center', fontWeight: 'bold', marginBottom: '10px', fontSize: '16px' }}>
                                        {field.tableName && (
                                            <td colSpan={cols} style={{ padding: '10px' }}>{field.tableName}</td>
                                        )}
                                    </tr>
                                    <tr>
                                        {Array.from({ length: cols }).map((_, c) => (
                                            <th key={`${field.internalName}-hdr-${c}`} style={{ border: '1px solid #ccc', padding: 8, backgroundColor: '#f0f0f0', textAlign: 'left' }}>{(colDefs[c] && colDefs[c].name) || `Column ${c + 1}`}</th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {Array.from({ length: rows }).map((_, r) => (
                                        <tr key={`${field.internalName}-r-${r}`}>
                                            {Array.from({ length: cols }).map((_, c) => (
                                                <td key={`${field.internalName}-cell-${r}-${c}`} style={{ border: '1px solid #ccc', padding: 8, minHeight: 30 }}>
                                                    {renderCell(r, c)}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                );
            case "htmlrender":
                // Render a fully prefilled HTML table saved in configuration
                const renderRows = field.tableRows || 0;
                const renderCols = field.tableColumns || 0;
                const renderHeaders = field.tableHeaders || [];
                const renderData: any[] = field.tableData || [];

                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <div style={{ border: "1px solid #000000", padding: "10px", borderRadius: "4px" }}>

                            <table style={{
                                borderCollapse: "collapse",
                                width: "100%",
                                marginTop: "10px",
                                border: "1px solid #ccc"
                            }}>
                                <thead>
                                    <tr style={{ textAlign: "center", fontWeight: "bold", marginBottom: "10px", fontSize: "16px" }}>
                                        {field.tableName && (
                                            <td colSpan={renderCols} style={{ padding: "10px" }}>{field.tableName}</td>
                                        )}
                                    </tr>
                                    <tr>
                                        {Array.from({ length: renderCols }).map((_, colIndex) => (
                                            <th key={`${field.internalName}-hdr-${colIndex}`} style={{
                                                border: "1px solid #ccc",
                                                padding: "8px",
                                                backgroundColor: "#f0f0f0",
                                                textAlign: "left"
                                            }}>{renderHeaders[colIndex] || `Column ${colIndex + 1}`}</th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {Array.from({ length: renderRows }).map((_, rowIndex) => (
                                        <tr key={`${field.internalName}-r-${rowIndex}`}>
                                            {Array.from({ length: renderCols }).map((_, colIndex) => {
                                                const colDef = (field.tableColumnDefs && field.tableColumnDefs[colIndex]) || { name: `Column ${colIndex + 1}`, type: 'text' };
                                                const cellValue = renderData[rowIndex] && typeof renderData[rowIndex][colIndex] !== 'undefined' ? renderData[rowIndex][colIndex] : "";

                                                // For yesno type, ensure "false" is displayed explicitly instead of blank
                                                let displayValue = cellValue;
                                                if (colDef.type === 'yesno' && !cellValue) {
                                                    displayValue = 'false';
                                                }

                                                return (
                                                    <td key={`${field.internalName}-cell-${rowIndex}-${colIndex}`} style={{
                                                        border: "1px solid #ccc",
                                                        padding: "8px",
                                                        minHeight: "30px",
                                                        backgroundColor: "white"
                                                    }}>
                                                        {displayValue}
                                                    </td>
                                                );
                                            })}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                );

            case "plainhtml":
                return (
                    <div key={field.internalName} style={fieldContainerStyle}>
                        <label style={labelStyle}>{field.label}</label>
                        <div style={{ marginTop: 8 }} dangerouslySetInnerHTML={{ __html: field.htmlContent || "" }} />
                    </div>
                );

            default:
                // Log unknown field types for debugging
                console.warn(`Unsupported field type: ${field.type}`);
                return null;
        }
    };

    /**
     * Render Method
     * Safely parses ConfigurationJSON and renders fields
     */
    public render(): React.ReactElement<AdminFormProps> {

        let parsedConfig = null;

        try {
            parsedConfig =
                typeof this.state.config?.ConfigurationJSON === "string"
                    ? JSON.parse(this.state.config.ConfigurationJSON)
                    : this.state.config?.ConfigurationJSON;
        } catch (error) {
            console.error("Invalid ConfigurationJSON format:", error);
        }

        const sortedFields = parsedConfig?.fields
            ? [...parsedConfig.fields].sort(
                (a: any, b: any) => a.order - b.order
            )
            : null;

        return (
            <div
                style={{
                    padding: "20px",
                    border: "1px solid #e5e5e5",
                    borderRadius: "8px",
                    boxShadow: "0 2px 6px rgba(0,0,0,0.08)"
                }}
            >
                <h2 style={{ marginBottom: "5px" }}>
                    {parsedConfig?.formTitle}
                </h2>

                <p style={{ marginBottom: "25px", color: "#666" }}>
                    {parsedConfig?.description}
                </p>

                {sortedFields ? (
                    sortedFields.map((field: any) =>
                        this.renderDynamicField(field)
                    )
                ) : (
                    <p>Loading...</p>
                )}

                <button
                    style={{
                        padding: "10px 18px",
                        backgroundColor: "#0078d4",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        marginTop: "10px"
                    }}
                    onClick={() => console.log(this.state.formData)}
                >
                    Submit
                </button>
            </div>
        );
    }
}