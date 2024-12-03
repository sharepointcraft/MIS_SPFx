import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './MisEventVersion.module.scss';

// Update the IMisEventVersionProps interface to include 'context' as a property
export interface IMisEventVersionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  userId: number; 
}

interface AttachmentLink {
  fileName: string;
  fileUrl: string;
  versionNumber: string;
}

interface IVersionInfo {
  version: string;
  plant: string;
  ndcCode: string;
  materialCode: number;
  description: string;
  product: string;
  packSize: number;
  conversionCost: number;
  rmc: number;
  pmc: number;
  consumables: number;
  acquisitionCostCMO: number;
  cop: number;
  freightDDPSea: number;
  updatedDate: string;
  remarksOnChanges: string;
  dosageForm: string;
  interestOnWc: number;
  cogs: number;
  strength: string;
  attachmentName?: string;
  attachmentUrl?: string;
}

export interface IMisEventVersionState {
  versionHistory: IVersionInfo[];
  loading: boolean;
  error: string | null;
  ndcCode: string;
  ndcCodeSuggestions: string[];
  allNdcCodes: string[];
  activeSuggestionIndex: number;
  showSuggestions: boolean;
  isMISGroupMember: boolean; // New state to track group membership
}

export default class MisEventVersion extends React.Component<IMisEventVersionProps, IMisEventVersionState> {
  constructor(props: IMisEventVersionProps) {
    super(props);

    const urlParams = new URLSearchParams(window.location.search);
    const ndcCodeFromUrl = urlParams.get('NDCCode') || '';

    this.state = {
      versionHistory: [],
      loading: true,
      error: null,
      ndcCode: ndcCodeFromUrl,
      ndcCodeSuggestions: [],
      allNdcCodes: [],
      activeSuggestionIndex: -1,
      showSuggestions: false,
      isMISGroupMember: false
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.checkMISGroupMembership();
    await this.fetchAllNdcCodes();

    if (this.state.ndcCode) {
      await this.getItemVersionHistory(this.state.ndcCode);
    } else {
      this.setState({ loading: false });
    }
  }

  private async checkMISGroupMembership(): Promise<void> {
    try {
      const groupUrl = `${this.props.siteUrl}/_api/web/sitegroups/getByName('MIS')/users?$filter=Id eq ${this.props.userId}`;
      const response = await this.props.spHttpClient.get(groupUrl, SPHttpClient.configurations.v1);
  
      if (response.ok) {
        const data = await response.json();
        const isMISGroupMember = data.value.length > 0;
        this.setState({ isMISGroupMember });
      } else {
        console.error('Failed to fetch group membership status');
      }
    } catch (error) {
      console.error('Error checking MIS group membership', error);
    }
  }
  

  private async fetchAllNdcCodes(): Promise<void> {
    try {
      const ndcCodeUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items?$select=NDCCode`;
      const response = await this.props.spHttpClient.get(ndcCodeUrl, SPHttpClient.configurations.v1);
      const data = await response.json();

      const allNdcCodes = data.value.map((item: any) => item.NDCCode);
      this.setState({ allNdcCodes });
    } catch (error) {
      console.error('Error fetching NDC codes', error);
      this.setState({ error: 'Error fetching NDC codes' });
    }
  }


  // Fetch version history of the list item and associated document set attachments
  private async getItemVersionHistory(ndcCode: string): Promise<void> {
    if (!ndcCode) {
      console.error('NDCCode not found in the input');
      this.setState({ loading: false, error: 'NDCCode not found in the input' });
      return;
    }
  
    try {
      const listItemId = await this.getListItemId(ndcCode);
      if (listItemId === -1) {
        this.setState({ loading: false, error: `No item found for NDCCode: ${ndcCode}` });
        return;
      }
  
      const versionHistoryUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items(${listItemId})/versions?$select=VersionLabel,Plant,NDCCode,Material_code,Description,Product,Pack_size,Conversion_cost,RMC,PMC,Consumables,Acquisition_Cost_CMO,COP,Freight_DDP_Sea,Updated_Date,Remarks_on_Changes,Dosage_form,Interest_on_Wc,COGS,Strength`;
      const response = await this.props.spHttpClient.get(versionHistoryUrl, SPHttpClient.configurations.v1);
      const versionHistoryData = await response.json();
  
      const attachments = await this.getAttachmentsFromDocSet(ndcCode);
  
      const convertVersionToComparable = (version: string): number => parseInt(version.split('.')[0], 10);
  
      const versionHistory: IVersionInfo[] = versionHistoryData.value.map((version: any) => {
        const comparableVersion = convertVersionToComparable(version.VersionLabel);
        const attachment = attachments.find(a => parseInt(a.versionNumber, 10) === comparableVersion);
  
        return {
          version: version.VersionLabel,
          plant: version.Plant,
          ndcCode: version.NDCCode,
          materialCode: version.Material_x005f_code,
          description: version.Description,
          product: version.Product,
          packSize: version.Pack_x005f_size,
          conversionCost: version.Conversion_x005f_cost,
          rmc: version.RMC,
          pmc: version.PMC,
          consumables: version.Consumables,
          acquisitionCostCMO: version.Acquisition_x005f_Cost_x005f_CMO,
          cop: version.COP,
          freightDDPSea: version.Freight_x005f_DDP_x005f_Sea,
          updatedDate: version.Updated_x005f_Date,
          remarksOnChanges: version.Remarks_x005f_on_x005f_Changes,
          dosageForm: version.Dosage_x005f_form,
          interestOnWc: version.Interest_x005f_on_x005f_Wc,
          cogs: version.COGS,
          strength: version.Strength,
          attachmentName: attachment ? attachment.fileName : undefined,
          attachmentUrl: attachment ? `${this.props.siteUrl}${attachment.fileUrl}` : undefined
        };
      });
  
      this.setState({ versionHistory, loading: false });
    } catch (error) {
      console.error('Error retrieving version history or attachments', error);
      this.setState({ loading: false, error: 'Error retrieving version history or attachments' });
    }
  }

  // Fetch the list item ID based on the NDCCode
  private async getListItemId(ndcCode: string): Promise<number> {
    const listUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items?$filter=NDCCode eq '${ndcCode}'`;
    const response = await this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
    const data = await response.json();

    if (data.value.length > 0) {
      return data.value[0].Id;
    }

    return -1; // Return -1 if no item is found
  }

  // Fetch attachments from the document set in the MIS_Attachment library based on the NDCCode
  private async getAttachmentsFromDocSet(ndcCode: string): Promise<AttachmentLink[]> {
    const attachments: AttachmentLink[] = [];
    try {
      const docSetUrl = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('MIS_Attachement/${ndcCode}')/Files?$expand=ListItemAllFields&$select=Name,ServerRelativeUrl,ListItemAllFields/Version_number,ListItemAllFields/ID,UIVersionLabel`;

      const response: SPHttpClientResponse = await this.props.spHttpClient.get(docSetUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const files = await response.json();

      files.value.forEach((file: any) => {
        const fileUrl = file.ServerRelativeUrl.replace('/sites/DevJay', '');

        attachments.push({
          fileName: file.Name,
          fileUrl: fileUrl,
          versionNumber: file.ListItemAllFields?.Version_number ?? 'N/A'
        });
      });
    } catch (error) {
      console.error('Error fetching document set attachments:', error);
      this.setState({ error: 'Error fetching document set attachments' });
    }
    return attachments;
  }

  // Handle the search input change and filter NDC codes for autocomplete
  private handleInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const userInput = event.target.value;
    const filteredSuggestions = this.state.allNdcCodes.filter(ndcCode =>
      ndcCode.toLowerCase().includes(userInput.toLowerCase())
    );

    this.setState({
      ndcCode: userInput,
      ndcCodeSuggestions: filteredSuggestions,
      activeSuggestionIndex: -1, // Reset active suggestion index
      showSuggestions: true // Show suggestions when user types
    });
  };

  // Handle keyboard events for navigation and selection of suggestions
  private handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    const { activeSuggestionIndex, ndcCodeSuggestions } = this.state;

    // User pressed the "Enter" key
    if (event.key === 'Enter') {
      if (activeSuggestionIndex >= 0 && activeSuggestionIndex < ndcCodeSuggestions.length) {
        // If a suggestion is active, select it
        this.handleSuggestionClick(ndcCodeSuggestions[activeSuggestionIndex]);
      } else {
        // Otherwise, trigger the search with the entered text
        this.handleSearch(event as any);
      }
    }

    // User pressed the "Arrow Up" key
    else if (event.key === 'ArrowUp') {
      if (activeSuggestionIndex === 0) {
        this.setState({ activeSuggestionIndex: ndcCodeSuggestions.length - 1 });
      } else {
        this.setState({ activeSuggestionIndex: activeSuggestionIndex - 1 });
      }
    }

    // User pressed the "Arrow Down" key
    else if (event.key === 'ArrowDown') {
      if (activeSuggestionIndex === ndcCodeSuggestions.length - 1) {
        this.setState({ activeSuggestionIndex: 0 });
      } else {
        this.setState({ activeSuggestionIndex: activeSuggestionIndex + 1 });
      }
    }
  };

  // Handle suggestion click (update the input with the selected NDC code)
  private handleSuggestionClick = (ndcCode: string): void => {
    this.setState(
      {
        ndcCode,
        ndcCodeSuggestions: [], // Clear suggestions after selection
        showSuggestions: false, // Hide suggestions after selection
        activeSuggestionIndex: -1 // Reset active suggestion index
      },
      () => {
        // Trigger search after selection
        this.getItemVersionHistory(ndcCode);
      }
    );
  };

  // Handle search when the form is submitted
  private handleSearch = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault(); // Prevent default form submission
    const { ndcCode } = this.state;

    // Trigger search based on NDC code
    this.getItemVersionHistory(ndcCode);
  };

  public render(): React.ReactElement<IMisEventVersionProps> {
    const { versionHistory, loading, error, ndcCode, ndcCodeSuggestions, activeSuggestionIndex, showSuggestions, isMISGroupMember } = this.state;


    return (
      <div className={styles.misEventVersion}>
        <form onSubmit={this.handleSearch}>
          <input
            type="text"
            value={ndcCode}
            onChange={this.handleInputChange}
            onKeyDown={this.handleKeyDown}
            placeholder="Search NDC Code"
            className={styles.searchInput}
          />
          {showSuggestions && ndcCodeSuggestions.length > 0 && (
            <ul className={styles.suggestionsList}>
              {ndcCodeSuggestions.map((suggestion, index) => (
                <li
                  key={suggestion}
                  className={index === activeSuggestionIndex ? styles.activeSuggestion : ''}
                  onClick={() => this.handleSuggestionClick(suggestion)}
                >
                  {suggestion}
                </li>
              ))}
            </ul>
          )}
          <button type="submit">Search</button>
        </form>

        {loading && <p>Loading...</p>}
        {error && <p className={styles.error}>{error}</p>}
        {versionHistory.length > 0 && (
        <div className={styles.tablecontainer}>
          <table className={styles.versionhistorytable}>
            <thead>
              <tr>
                <th>Version</th>
                {isMISGroupMember && <th>Attachment</th>}
                <th>NDC Code</th>
                <th>Plant</th>
                <th>Dosage Form</th>
                <th>Material Code</th>
                <th>Description</th>
                <th>Product</th>
                <th>Strength</th>
                <th>Pack Size</th>
                <th>RMC</th>
                <th>PMC</th>
                <th>Consumables</th>
                <th>Conversion Cost</th>
                <th>Acquisition Cost CMO</th>
                <th>Interest on WC</th>
                <th>COP</th>
                <th>Freight DDP Sea</th>
                <th>COGS</th>
                <th>Updated Date</th>
                <th>Remarks on Changes</th>                
              </tr>
            </thead>
            <tbody>
              {versionHistory.map((version, index) => (
                <tr key={index}>
                  <td>{version.version}</td>
                  {isMISGroupMember && (
                    <td>
                      {version.attachmentName ? (
                        <a href={`${version.attachmentUrl}?web=1`} target="_blank" rel="noopener noreferrer">
                          {version.attachmentName}
                        </a>
                      ) : (
                        'No attachment'
                      )}
                    </td>
                  )}
                  <td>{version.ndcCode}</td>
                  <td>{version.plant}</td>
                  <td>{version.dosageForm}</td>
                  <td>{version.materialCode}</td>
                  <td>{version.description}</td>
                  <td>{version.product}</td>
                  <td>{version.strength}</td>
                  <td>{version.packSize}</td>                 
                  <td>{version.rmc}</td>
                  <td>{version.pmc}</td>
                  <td>{version.consumables}</td>
                  <td>{version.conversionCost}</td>
                  <td>{version.acquisitionCostCMO}</td>
                  <td>{version.interestOnWc}</td>
                  <td>{version.cop}</td>
                  <td>{version.freightDDPSea}</td>
                  <td>{version.cogs}</td>
                  <td>
                    {version.updatedDate
                      ? new Date(version.updatedDate).toLocaleDateString('en-US', {
                          month: '2-digit',
                          day: '2-digit',
                          year: 'numeric'
                        })
                      : ''}
                  </td>
                  <td>{version.remarksOnChanges}</td>                  
                </tr>
              ))}
            </tbody>
          </table>
        </div>  
        )}
      </div>
    );
  }
}