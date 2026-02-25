import * as React from 'react';
import { SPFI } from "@pnp/sp";
import { Dialog } from "@fluentui/react/lib/Dialog";
import Eform from './EForm';
import NewForm from './NewEForm';
import styles from './AdminWebPart.module.scss';

/**
 * Props Interface
 * sp: Injected SPFI instance for SharePoint data operations
 */
interface IAdminWEbPartProps {
  sp: SPFI;
}

/**
 * State Interface
 * config: Stores parsed ConfigList items
 * isViewDialogOpen: Controls view dialog visibility
 * isEditDialogOpen: Controls edit dialog visibility
 * isNewFormDialogOpen: Controls new form dialog visibility
 * selectedItem: Stores currently selected form item
 */
interface IAdminWEbPartState {
  config: any[];
  isViewDialogOpen: boolean;
  isEditDialogOpen: boolean;
  isNewFormDialogOpen: boolean;
  selectedItem?: any;
}

/**
 * Admin WebPart Component
 * 
 * Responsible for:
 * - Loading configuration forms from SharePoint ConfigList
 * - Rendering list of forms
 * - Managing dialog states (View / Edit / New)
 * - Refreshing data post form creation
 */
export default class AdminWEbPart extends React.Component<IAdminWEbPartProps, IAdminWEbPartState> {

  constructor(props: IAdminWEbPartProps) {
    super(props);

    // Initial component state
    this.state = {
      config: [],
      isViewDialogOpen: false,
      isEditDialogOpen: false,
      isNewFormDialogOpen: false,
      selectedItem: undefined
    };
  }

  /**
   * Opens "New Form" dialog
   */
  private _onNewFormClick = (): void => {
    this.setState({
      isNewFormDialogOpen: true
    });
  };

  /**
   * Opens View dialog for selected form
   */
  private _onViewClick = (item: any): void => {
    this.setState({
      isViewDialogOpen: true,
      selectedItem: item
    });
  };

  /**
   * Closes all dialogs
   * Centralized dialog state reset
   */
  private _closeDialogs = (): void => {
    this.setState({
      isViewDialogOpen: false,
      isEditDialogOpen: false,
      isNewFormDialogOpen: false
    });
  };

  /**
   * Lifecycle Method
   * Triggered after component mounts
   * Loads configuration items from SharePoint
   */
  public async componentDidMount(): Promise<void> {
    await this._loadItems();
  }

  /**
   * Loads items from ConfigList
   * - Fetches raw list items
   * - Parses ConfigurationJSON column
   * - Extracts formTitle & description
   * - Sorts alphabetically (case-insensitive)
   */
  private async _loadItems(): Promise<void> {
    try {

      const items = await this.props.sp.web.lists
        .getByTitle("ConfigList")
        .items.filter("Status eq 'Active'")() ;

      // Parse JSON configuration for each item
      const parsedItems = items.map(item => {

        // Safely parse ConfigurationJSON
        const json = item.ConfigurationJSON
          ? JSON.parse(item.ConfigurationJSON)
          : null;

        return {
          Id: item.Id,
          formTitle: json?.formTitle || "Untitled Form",
          description: json?.description || ""
        };
      });

      // Sort forms alphabetically (case-insensitive)
      parsedItems.sort((a, b) =>
        a.formTitle.localeCompare(b.formTitle, undefined, {
          sensitivity: "base"
        })
      );

      // Update state with processed configuration
      this.setState({ config: parsedItems });

    } catch (error) {

      // Production Logging
      console.error("Error fetching ConfigList items:", error);
    }
  }

  /**
   * Refreshes configuration list
   * Used after creating a new form
   */
  private _refreshData = async (): Promise<void> => {
    await this._loadItems();
  };

  /**
   * Main Render Method
   * Renders:
   * - Create New Form link
   * - List of configured forms
   * - View Dialog
   * - Edit Dialog
   * - New Form Dialog
   */
  public render(): React.ReactElement<IAdminWEbPartProps> {

    return (
      <div style={{ padding: "20px" }}>

        {/* Create New Form Trigger */}
        <div style={{ textAlign: "left", marginBottom: "15px" }}>
          <a
            href="#"
            onClick={(e) => {
              e.preventDefault();
              this._onNewFormClick();
            }}
            style={{
              fontSize: "18px",
              fontWeight: 600,
              color: "#0078d4",
              textDecoration: "none",
              cursor: "pointer"
            }}
          >
            + Create New Form
          </a>
        </div>

        {/* Section Header */}
        <h2 style={{
          marginBottom: "20px",
          backgroundColor: 'grey',
          color: 'white',
          padding: '10px'
        }}>
          Eforms
        </h2>

        {/* Config List Rendering */}
        <ul style={{
          listStyleType: "none",
          padding: 0,
          margin: 0
        }}>

          {this.state.config.map((item) => (

            <li
              key={item.Id}
              style={{
                padding: "12px 8px",
                borderBottom: "1px solid #eaeaea"
              }}
            >

              {/* Form Title Hyperlink */}
              <a
                href="#"
                onClick={(e) => {
                  e.preventDefault();
                  this._onViewClick(item);
                }}
                style={{
                  fontSize: "16px",
                  fontWeight: 500,
                  color: "#0078d4",
                  textDecoration: "none",
                  cursor: "pointer"
                }}
              >
                {item.formTitle}
              </a>

            </li>

          ))}

        </ul>

        {/* View Form Dialog */}
        <Dialog
          modalProps={{ className: styles.narrowDialog }}
          hidden={!this.state.isViewDialogOpen}
          onDismiss={this._closeDialogs}
        >
          <Eform
            formIDValue={this.state.selectedItem?.Id}
            sp={this.props.sp}
          />
        </Dialog>

        {/* Placeholder Edit Dialog (Future Enhancement) */}
        <Dialog
          maxWidth={850}
          hidden={!this.state.isEditDialogOpen}
          onDismiss={this._closeDialogs}
        >
          <div>Edit Form</div>
        </Dialog>

        {/* New Form Creation Dialog */}
        <Dialog
          modalProps={{ className: styles.narrowDialog }}
          hidden={!this.state.isNewFormDialogOpen}
          onDismiss={() => this.setState({ isNewFormDialogOpen: false })}
        >
          <NewForm
            sp={this.props.sp}
            onSuccess={() => {
              this._refreshData();   // Reload list after successful creation
              this._closeDialogs();  // Close dialog
            }}
          />
        </Dialog>

      </div>
    );
  }
}
