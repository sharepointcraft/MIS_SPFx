import * as React from 'react';
import { parse, ParseResult } from 'papaparse'; // For CSV parsing
import * as XLSX from 'xlsx'; // For Excel parsing
import { sp } from '@pnp/sp/presets/all'; // Ensure PnPjs is configured
import type { IMisPnpUoloadProps } from './IMisPnpUoloadProps';

interface ITableData {
  'NDC Code': string;
  Plant: string;
  'Dosage form': string;
}

interface IAttachment {
  file: File;
  ndcCode: string;
}

export default class MisPnpUpload extends React.Component<
  IMisPnpUoloadProps,
  { filePickerResult: File[]; tableData: ITableData[]; attachments: IAttachment[] }
> {
  constructor(props: IMisPnpUoloadProps) {
    super(props);
    this.state = {
      filePickerResult: [],
      tableData: [],
      attachments: [],
    };
  }

  public render(): React.ReactElement<IMisPnpUoloadProps> {
    return (
      <div>
        <button onClick={this._triggerFileInput}>Choose File</button>
        <input
          type="file"
          ref={(input) => (this.fileInput = input)}
          onChange={this._onFileSelected}
          style={{ display: 'none' }}
        />

        {this.state.tableData.length > 0 && this._renderTable()}

        <div style={{ marginTop: '10px' }}>
          <button onClick={this._handleSubmit} style={{ marginRight: '10px' }}>Submit</button>
          <button onClick={this._handleCancel}>Cancel</button>
        </div>
      </div>
    );
  }

  private fileInput: HTMLInputElement | null = null;

  // This method triggers the hidden file input element to open the file dialog directly.
  private _triggerFileInput = () => {
    if (this.fileInput) {
      this.fileInput.click();
    }
  };

  // Handle file selection
  private _onFileSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      const fileName = file.name?.toLowerCase();
      if (fileName.endsWith('.csv')) {
        const fileContent = await file.text();
        this._parseCSV(fileContent);
      } else if (fileName.endsWith('.xlsx')) {
        const arrayBuffer = await file.arrayBuffer();
        this._parseExcel(arrayBuffer);
      }
      this.setState({ filePickerResult: [file] });
    }
  };

  private _parseCSV = (csvContent: string) => {
    parse(csvContent, {
      header: true, // CSV contains header row
      skipEmptyLines: true,
      complete: (results: ParseResult<ITableData>) => {
        const rawData = results.data;
        this.setState({ tableData: rawData });
      },
      error: (error: any) => {
        console.error('Error parsing CSV', error);
      },
    });
  };

  private _parseExcel = (arrayBuffer: ArrayBuffer) => {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const rawData = json.slice(3); // Assuming data starts from row 4

    const data: ITableData[] = rawData.map((row: any) => ({
      'NDC Code': row[0],
      Plant: row[1],
      'Dosage form': row[2],
    }));

    this.setState({ tableData: data });
  };

  private _renderTable = () => {
    const { tableData } = this.state;

    if (tableData.length === 0) return null;

    const headers = ['NDC Code', 'Plant', 'Dosage form'];

    return (
      <table>
        <thead>
          <tr>
            {headers.map((header, index) => (
              <th key={index}>{header}</th>
            ))}
            <th>Attachment</th>
          </tr>
        </thead>
        <tbody>
          {tableData.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {headers.map((header, colIndex) => (
                <td key={colIndex}>{row[header as keyof ITableData]}</td>
              ))}
              <td>
                <input
                  type="file"
                  onChange={(e) => this._handleFileChange(e, row['NDC Code'])}
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  private _handleFileChange = (event: React.ChangeEvent<HTMLInputElement>, ndcCode: string) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const attachments = [...this.state.attachments];
      attachments.push({ file: files[0], ndcCode });
      this.setState({ attachments });
    }
  };

  private _handleCancel = () => {
    // Clear the selected files and the grid view data
    this.setState({
      filePickerResult: [],
      tableData: [],
      attachments: [],
    });
  };

  private _handleSubmit = async () => {
    const { tableData, attachments } = this.state;

    if (tableData.length === 0) {
      console.error('No table data to save');
      return;
    }

    try {
      // Save all grid view data to the SharePoint list
      for (const row of tableData) {
        const ndcCode = row['NDC Code']; // Using display name from grid
        const plant = row['Plant'];
        const dosageForm = row['Dosage form'];

        // Check if an item with the same NDC Code already exists
        const existingItems = await sp.web.lists.getByTitle('MIS_Upload_File')
          .items.filter(`NDCCode eq '${ndcCode}'`) // Use the correct internal column name 'NDCCode'
          .top(1) // We only care about one existing item
          .get();

        if (existingItems.length > 0) {
          // If the item exists, update the version history and the record
          const existingItem = existingItems[0];
          const currentVersion = existingItem.VersionHistroy || 'V1'; // Use 'V1' as default if no version history exists
          const versionNumber = parseInt(currentVersion.substring(1)); // Extract the number from the version (e.g., V1 -> 1)
          const newVersion = `V${versionNumber + 1}`; // Increment version

          // Update existing record with new values and updated version
          await sp.web.lists.getByTitle('MIS_Upload_File').items.getById(existingItem.Id).update({
            Plant: plant,
            Dosage_x0020_form: dosageForm, // Correct internal name for 'Dosage form'
            VersionHistroy: newVersion, // Update 'VersionHistroy' column with the new version
          });

          console.log(`Record with NDC Code '${ndcCode}' updated to version ${newVersion}.`);
        } else {
          // If the item doesn't exist, create a new record with version history 'V1'
          await sp.web.lists.getByTitle('MIS_Upload_File').items.add({
            NDCCode: ndcCode, // Correct internal name for 'NDCCode'
            Plant: plant,
            Dosage_x0020_form: dosageForm, // Correct internal name for 'Dosage form'
            VersionHistroy: 'V1', // Start version history in 'VersionHistroy' column
          });

          console.log(`New record with NDC Code '${ndcCode}' created with version V1.`);
        }
      }

      // Handle file attachments similarly
      const attachmentsByNdcCode = attachments.reduce((acc: Record<string, IAttachment[]>, attachment) => {
        if (!acc[attachment.ndcCode]) {
          acc[attachment.ndcCode] = [];
        }
        acc[attachment.ndcCode].push(attachment);
        return acc;
      }, {});

      // Save attachments to the SharePoint document library
      for (const [ndcCode, attachments] of Object.entries(attachmentsByNdcCode)) {
        let documentSetId;

        try {
          const existingDocumentSet = await sp.web.lists
            .getByTitle('MIS_Attachment')
            .items.filter(`Title eq '${ndcCode}'`)
            .top(1)
            .get();

          if (existingDocumentSet.length > 0) {
            documentSetId = existingDocumentSet[0].Id;
          } else {
            const documentSet = await sp.web.lists.getByTitle('MIS_Attachment').items.add({
              ContentTypeId: '0x0120D52000A9F6E84B73F44EADDEADB84D61',
              Title: ndcCode,
            });
            documentSetId = documentSet.data.Id;
          }
        } catch (error) {
          console.error(`Error creating/retrieving document set for NDC Code ${ndcCode}`, error);
        }

        const folderUrl = `/sites/DevJay/MIS_Attachment/${documentSetId}`;
        const folderExists = await sp.web.getFolderByServerRelativeUrl(folderUrl).select('Exists').get()
          .then(() => true)
          .catch(() => false);

        if (!folderExists) {
          await sp.web.folders.add(folderUrl);
        }

        for (const attachment of attachments) {
          const { file } = attachment;
          await sp.web
            .getFolderByServerRelativeUrl(folderUrl)
            .files.add(file.name, file, true)
            .then(() => console.log(`File uploaded to ${folderUrl}`));
        }

        await sp.web.lists.getByTitle('MIS_Attachment').items.getById(documentSetId).update({
          NDCCode: ndcCode,
        });
      }

      console.log('All data and attachments have been successfully saved.');
    } catch (error) {
      console.error('Error saving data to SharePoint', error);
    }
  };




}
