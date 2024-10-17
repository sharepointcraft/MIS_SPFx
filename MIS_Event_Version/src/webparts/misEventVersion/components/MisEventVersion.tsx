import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IMisEventVersionProps } from './IMisEventVersionProps';
import styles from './MisEventVersion.module.scss';

interface AttachmentLink {
  fileName: string;
  fileUrl: string;
  versionNumber: string; // Custom 'Version_number' column value
}

interface IVersionInfo {
  version: string;
  attachmentName?: string; // Store the document name instead of the link
  attachmentUrl?: string; // Store the URL of the attachment
}

export interface IMisEventVersionState {
  versionHistory: IVersionInfo[];
  loading: boolean;
  error: string | null;
  ndcCode: string; // Store the user input NDCCode
  ndcCodeSuggestions: string[]; // Store the autocomplete suggestions
  allNdcCodes: string[]; // Store all NDC codes for autocomplete
  activeSuggestionIndex: number; // To track the active suggestion with arrow keys
  showSuggestions: boolean; // To control when to show suggestions
}

export default class MisEventVersion extends React.Component<IMisEventVersionProps, IMisEventVersionState> {
  constructor(props: IMisEventVersionProps) {
    super(props);

    // Get the NDCCode from URL parameters (if it exists)
    const urlParams = new URLSearchParams(window.location.search);
    const ndcCodeFromUrl = urlParams.get('ndcCode') || ''; // Get the ndcCode or use an empty string

    this.state = {
      versionHistory: [],
      loading: true, // Set loading to true initially while fetching data
      error: null,
      ndcCode: ndcCodeFromUrl, // Initialize with NDC code from URL
      ndcCodeSuggestions: [], // Autocomplete suggestions array
      allNdcCodes: [], // Array to store all NDC codes
      activeSuggestionIndex: -1, // No active suggestion by default
      showSuggestions: false // Suggestions will only show when user types
    };
  }

  public async componentDidMount(): Promise<void> {
    // Fetch all NDC codes for the autocomplete
    await this.fetchAllNdcCodes();

    // If there is an NDC code in the URL, fetch the version history
    if (this.state.ndcCode) {
      await this.getItemVersionHistory(this.state.ndcCode);
    } else {
      this.setState({ loading: false }); // No NDC code found, stop loading
    }
  }

  // Fetch all NDC codes for the autocomplete suggestions
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
      // Fetch the list item ID using NDCCode
      const listItemId = await this.getListItemId(ndcCode);
      if (listItemId === -1) {
        this.setState({ loading: false, error: `No item found for NDCCode: ${ndcCode}` });
        return;
      }

      // Fetch version history of the list item
      const versionHistoryUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items(${listItemId})/versions`;
      const response = await this.props.spHttpClient.get(versionHistoryUrl, SPHttpClient.configurations.v1);
      const versionHistoryData = await response.json();

      // Fetch the attachments for the document set from MIS_Attachment library
      const attachments = await this.getAttachmentsFromDocSet(ndcCode);

      const convertVersionToComparable = (version: string): number => {
        return parseInt(version.split('.')[0], 10); // Get only the major version number
      };

      const versionHistory: IVersionInfo[] = versionHistoryData.value.map((version: any) => {
        const comparableVersion = convertVersionToComparable(version.VersionLabel);
        const attachment = attachments.find(a => parseInt(a.versionNumber, 10) === comparableVersion);

        return {
          version: version.VersionLabel,
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
    const { versionHistory, loading, error, ndcCode, ndcCodeSuggestions, activeSuggestionIndex, showSuggestions } = this.state;

    return (
      <div className={styles.misEventVersion}>
        <form onSubmit={this.handleSearch}>
          <input
            type="text"
            value={ndcCode}
            onChange={this.handleInputChange}
            onKeyDown={this.handleKeyDown} // Add keydown handler for keyboard navigation
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
          <table className={styles.versionTable}>
            <thead>
              <tr>
                <th>Version</th>
                <th>Attachment</th>
              </tr>
            </thead>
            <tbody>
              {versionHistory.map((version, index) => (
                <tr key={index}>
                  <td>{version.version}</td>
                  <td>
                    {version.attachmentName ? (
                      <a href={version.attachmentUrl} target="_blank" rel="noopener noreferrer">
                        {version.attachmentName}
                      </a>
                    ) : (
                      'No attachment'
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    );
  }
}
