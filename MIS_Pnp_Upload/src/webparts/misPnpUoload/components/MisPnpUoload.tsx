import * as React from 'react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';
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
  index: number;
}

export default class MisPnpUoload extends React.Component<
  IMisPnpUoloadProps,
  { filePickerResult: IFilePickerResult[]; tableData: ITableData[]; attachments: IAttachment[] }
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
        <FilePicker
          bingAPIKey="<BING API KEY>" // Replace with your actual Bing API Key or remove if not needed
          accepts={['.csv', '.xlsx']}
          buttonIcon="FileImage"
          onSave={this._onFilePickerSave}
          onChange={this._onFilePickerChange}
          context={this.props.context} // Ensure context is passed
        />
        {this.state.tableData.length > 0 && this._renderTable()}
        <button onClick={this._handleSubmit}>Submit</button>
      </div>
    );
  }

  private _onFilePickerChange = (filePickerResult: IFilePickerResult[]) => {
    this.setState({ filePickerResult });
  };

  private _onFilePickerSave = async (filePickerResult: IFilePickerResult[]) => {
    this.setState({ filePickerResult });
    if (filePickerResult && filePickerResult.length > 0) {
      for (const item of filePickerResult) {
        try {
          const fileResultContent = await item.downloadFileContent();
          const fileName = item.fileName?.toLowerCase();

          if (fileName?.endsWith('.csv')) {
            const fileContent = await fileResultContent.text();
            this._parseCSV(fileContent);
          } else if (fileName?.endsWith('.xlsx')) {
            const arrayBuffer = await fileResultContent.arrayBuffer();
            this._parseExcel(arrayBuffer);
          }
        } catch (error) {
          console.error('Error processing file content', error);
        }
      }
    }
  };

  private _parseCSV = (csvContent: string) => {
    parse(csvContent, {
      header: true, // CSV contains header row
      skipEmptyLines: true,
      complete: (results: ParseResult<ITableData>) => {
        const rawData = results.data; // Start reading from row 1 (since it's CSV, header is included)
        console.log('Parsed CSV Data:', rawData); // Debugging line
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

    const rawData = json.slice(3); // Data starts from row 5

    const data: ITableData[] = rawData.map((row: any) => {
      const rowData: ITableData = {
        'NDC Code': row[0], // Adjust indices based on your actual Excel file layout
        Plant: row[1],
        'Dosage form': row[2]
      };
      return rowData;
    });

    console.log('Parsed Excel Data:', data); // Debugging line
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
                <td key={colIndex}>{this._formatCell(row[header as keyof ITableData], header)}</td>
              ))}
              <td>
                <input
                  type="file"
                  onChange={(e) => this._handleFileChange(e, rowIndex)}
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  private _handleFileChange = (event: React.ChangeEvent<HTMLInputElement>, rowIndex: number) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const attachments = [...this.state.attachments];
      attachments[rowIndex] = { file: files[0], index: rowIndex };
      this.setState({ attachments });
    }
  };

  private _handleSubmit = async () => {
    const { tableData, attachments } = this.state;

    if (tableData.length === 0) {
      console.error('No table data to save');
      return;
    }

    try {
      for (let i = 0; i < tableData.length; i++) {
        const row = tableData[i];

        if (row['NDC Code'] && row['Plant'] && row['Dosage form']) {
          // Add data to the 'MIS_Upload_File' list
          await sp.web.lists.getByTitle('MIS_Upload_File').items.add({
            NDCCode: row['NDC Code'],
            Plant: row['Plant'],
            Dosage_x0020_form: row['Dosage form']
          });

          const ndcCode = row['NDC Code'];

          // Check if there is an attachment for this row
          if (attachments[i]) {
            const { file } = attachments[i];
            
            // Create a folder in 'MIS_Attachment' named after the NDC Code
            const folderUrl = `/sites/DevJay/MIS_Attachment/${ndcCode}`;
            await sp.web.folders.add(folderUrl);

            // Add the attachment to the corresponding folder
            await sp.web.getFolderByServerRelativeUrl(folderUrl).files.add(file.name, file, true);
          }
        } else {
          console.warn(`Row ${i + 1} has missing data. Skipping row.`);
        }
      }

      console.log('All data and attachments have been successfully saved.');
    } catch (error) {
      console.error('Error saving data to SharePoint', error);
    }
  };

  private _formatCell = (value: any, header: string) => {
    // Example: Format cells for columns that need a dollar sign
    if (header === 'Price' || header === 'Cost') {
      return `$${value}`;
    }
    return value;
  };
}
