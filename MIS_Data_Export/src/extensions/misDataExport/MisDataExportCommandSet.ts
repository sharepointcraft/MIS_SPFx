import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp/presets/all';
import { Dialog } from '@microsoft/sp-dialog';
import * as XLSX from 'xlsx'; // Import the xlsx library

export interface IMisDataExportCommandSetProperties {
  sampleText: string; // You can use this for customization if needed
}

const LOG_SOURCE: string = 'MisDataExportCommandSet';

export default class MisDataExportCommandSet extends BaseListViewCommandSet<IMisDataExportCommandSetProperties> {

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MisDataExportCommandSet');

    // Conditional check for "DevJay" site and "MIS_Upload_File" list
    const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
    const currentListTitle = this.context.pageContext.list?.title; // Use optional chaining to handle undefined list

    if (currentSiteUrl.includes('DevJay') && currentListTitle === 'MIS_Upload_File') {
      // Allow the command to appear if the conditions match
      return Promise.resolve();
    } else {
      // Hide the command if the site or list does not match
      this._hideCommandBarButton();
      return Promise.resolve();
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'ExportExcel':
        this._exportToExcel();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  // Helper function to hide the command bar button
  private _hideCommandBarButton(): void {
    const exportCommand = this.tryGetCommand('ExportExcel');
    if (exportCommand) {
      exportCommand.visible = false;
    }
  }

  private async _exportToExcel(): Promise<void> {
    const listTitle = this.context.pageContext.list?.title; // Safely check for list context

    if (!listTitle) {
      Dialog.alert('No list context available.');
      return;
    }

    try {
      // Fetch all list items
      const items: any[] = await sp.web.lists.getByTitle(listTitle).items.top(5000)(); // Adjust the top count as needed

      if (items.length === 0) {
        Dialog.alert('No data available to export.');
        return;
      }

      // Map internal field names to display names
      const fieldMapping = {
        "NDCCode": "NDC Code",
        "Plant": "Plant",
        "Dosage_form": "Dosage form",
        "Material_code": "Material code",
        "Description": "Description",
        "Product": "Product",
        "Strength": "Strength",
        "Pack_size": "Pack size",
        "Conversion_cost": "Conversion cost",
        "RMC": "RMC",
        "PMC": "PMC",
        "Consumables": "Consumables",
        "Acquisition_Cost_CMO": "Acquisition Cost CMO",
        "Interest_on_Wc": "Interest on Wc",
        "COP": "COP",
        "Freight_DDP_Sea": "Freight DDP Sea",
        "COGS": "COGS",
        "Updated_Date": "Updated Date",
        "Remarks_on_Changes": "Remarks on Changes"
      };

      // List of columns where the "$" sign should be added
      const currencyFields = [
        "Conversion cost",
        "RMC",
        "PMC",
        "Consumables",
        "Acquisition Cost CMO",
        "Interest on Wc",
        "COP",
        "Freight DDP Sea",
        "COGS"
      ];

      // Filter and rename items
      const filteredItems = items.map(item => {
        const filteredItem: { [key: string]: any } = {};
        (Object.keys(fieldMapping) as Array<keyof typeof fieldMapping>).forEach(internalField => {
          let value = item[internalField];

          // Format the "Updated_Date" field as "MM-DD-YYYY"
          if (internalField === "Updated_Date" && value) {
            const dateValue = new Date(value);
            filteredItem[fieldMapping[internalField]] = dateValue.toLocaleDateString('en-US', {
              year: 'numeric',
              month: '2-digit',
              day: '2-digit',
            });
          } else if (currencyFields.includes(fieldMapping[internalField]) && value) {
            // If the field is a currency field, prepend the $ sign
            filteredItem[fieldMapping[internalField]] = `$${value}`;
          } else {
            filteredItem[fieldMapping[internalField]] = value; // Apply display name
          }
        });
        return filteredItem;
      });

      // Convert the filtered data to XLSX format and trigger the download
      this._downloadExcel(filteredItems, `${listTitle}.xlsx`);

      // Log the export action after the export is successful
      await this._logExportAction();
    } catch (error) {
      Dialog.alert('Error exporting list data to Excel: ' + error.message);
    }
  }

  private _downloadExcel(data: any[], filename: string): void {
    // Create a new workbook and a new worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);

    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Generate a file and trigger the download
    XLSX.writeFile(wb, filename);
  }

  // Method to log the export action to the "MIS_Version_Logs" list
  private async _logExportAction(): Promise<void> {
    const userName = this.context.pageContext.user.displayName;
    const userEmail = this.context.pageContext.user.email;

    try {
      const userId = await this._getUserId(userEmail); // Get the user ID based on email

      if (!userId) {
        throw new Error('Failed to retrieve user ID.');
      }

      // Log entry in the "MIS_Version_Logs" list
      const listTitle = 'MIS_Version_Logs';
      const logEntry = {
        Title: `Export to Excel by ${userName}`,
        Log_CreatorId: userId, // Use the user ID for the People Picker column
      };

      await sp.web.lists.getByTitle(listTitle).items.add(logEntry);
      Log.info(LOG_SOURCE, `Export action logged successfully for ${userName}`);
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to log export action: ${error.message}`));
    }
  }

  // Helper function to get the user's ID based on their email
  private async _getUserId(email: string): Promise<number | null> {
    try {
      const result = await sp.web.siteUsers.getByEmail(email).get();
      return result.Id;
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to get user ID for email: ${email}`));
      return null;
    }
  }
}
