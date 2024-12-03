import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';
import { sp } from '@pnp/sp/presets/all';
import { Dialog } from '@microsoft/sp-dialog';
import * as XLSX from 'xlsx';

export interface IMisDataExportCommandSetProperties {
  sampleText: string; // For customization if needed
}

const LOG_SOURCE: string = 'MisDataExportCommandSet';

export default class MisDataExportCommandSet extends BaseListViewCommandSet<IMisDataExportCommandSetProperties> {

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MisDataExportCommandSet');

    const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
    const currentListTitle = this.context.pageContext.list?.title;

    // Show Export button only if conditions are met
    if (currentSiteUrl.includes('DevJay') && currentListTitle === 'MIS_Upload_File') {
      this.tryGetCommand('ExportExcel').visible = true;
      Log.info(LOG_SOURCE, 'Export button set to visible.');
    } else {
      this._hideCommandBarButton();
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'ExportExcel':
        Log.info(LOG_SOURCE, 'Export to Excel button clicked.');
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
      Log.info(LOG_SOURCE, 'Export button hidden as it does not match conditions.');
    }
  }

  private async _exportToExcel(): Promise<void> {
    const listTitle = this.context.pageContext.list?.title;
    if (!listTitle) {
      Dialog.alert('No list context available.');
      return;
    }

    try {
      // Fetch all list items
      const items: any[] = await sp.web.lists.getByTitle(listTitle).items.top(5000)();
      if (items.length === 0) {
        Dialog.alert('No data available to export.');
        return;
      }

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

      const currencyFields = ["Conversion cost", "RMC", "PMC", "Consumables", "Acquisition Cost CMO", "Interest on Wc", "COP", "Freight DDP Sea", "COGS"];

      const filteredItems = items.map(item => {
        const filteredItem: { [key: string]: any } = {};
        
        for (const internalField in fieldMapping) {
          const displayName = fieldMapping[internalField as keyof typeof fieldMapping];
          
          if (displayName) {
            let value = item[internalField];
            
            // Format date and currency fields
            if (internalField === "Updated_Date" && value) {
              const dateValue = new Date(value);
              filteredItem[displayName] = dateValue.toLocaleDateString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
              });
            } else if (currencyFields.includes(displayName) && value) {
              filteredItem[displayName] = `$${value}`;
            } else {
              filteredItem[displayName] = value;
            }
          }
        }
        
        return filteredItem;
      });

      // Trigger Excel download
      this._downloadExcel(filteredItems, `${listTitle}.xlsx`);
      
      // Log export action after successful download
      await this._logExportAction();
      
    } catch (error) {
      Dialog.alert('Error exporting list data to Excel: ' + error.message);
      Log.error(LOG_SOURCE, new Error('Export error: ' + error.message));
    }
  }

  private _downloadExcel(data: any[], filename: string): void {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename);
  }

  private async _logExportAction(): Promise<void> {
    const userName = this.context.pageContext.user.displayName;
    const userEmail = this.context.pageContext.user.email;

    try {
      const userId = await this._getUserId(userEmail);
      if (!userId) throw new Error('Failed to retrieve user ID.');

      const logEntry = {
        Title: `Export to Excel by ${userName}`,
        Log_CreatorId: userId,
      };

      await sp.web.lists.getByTitle('MIS_Version_Logs').items.add(logEntry);
      Log.info(LOG_SOURCE, `Export action logged for ${userName}`);
    } catch (error) {
      Log.error(LOG_SOURCE, new Error('Failed to log export action: ' + error.message));
    }
  }

  private async _getUserId(email: string): Promise<number | null> {
    try {
      const result = await sp.web.siteUsers.getByEmail(email).get();
      return result.Id;
    } catch (error) {
      Log.error(LOG_SOURCE, new Error('User ID retrieval error for email: ' + email));
      return null;
    }
  }
}
