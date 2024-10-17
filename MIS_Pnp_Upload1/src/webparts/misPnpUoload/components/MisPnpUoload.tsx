import * as React from "react";
import { parse, ParseResult } from "papaparse";
import * as XLSX from "xlsx";
import { sp } from "@pnp/sp/presets/all";
import type { IMisPnpUoloadProps } from "./IMisPnpUoloadProps";
import "./MisPnpUoload.module.scss";

interface ITableData {
  "NDC Code": string;
  Plant: string;
  Dosage_form: string;
  Material_code: number;
  Description: string;
  Product: string;
  Strength: number;
  Pack_size: number;
  RMC: string;
  PMC: string;
  Consumables: string;
  Conversion_cost: string;
  Acquisition_Cost_CMO: string;
  Interest_on_Wc: string;
  COP: string;
  Freight_DDP_Sea: string;
  COGS: string;
  Updated_Date: string;
  Remarks_on_Changes: string;
}

interface IAttachment {
  file: File;
  ndcCode: string;
}

export default class MisPnpUpload extends React.Component<
  IMisPnpUoloadProps,
  {
    filePickerResult: File[];
    tableData: ITableData[];
    attachments: IAttachment[];
    fileName: string;
    isSubmitDisabled: boolean;
    loading: boolean; // New loading state
  }
> {
  constructor(props: IMisPnpUoloadProps) {
    super(props);
    this.state = {
      filePickerResult: [],
      tableData: [],
      attachments: [],
      fileName: "",
      isSubmitDisabled: false,
      loading: false, // Initialize loading as false
    };
  }

  componentDidMount() {
    const isSubmitDisabled =
      localStorage.getItem("isSubmitDisabled") === "true";
    this.setState({ isSubmitDisabled });
  }

  public render(): React.ReactElement<IMisPnpUoloadProps> {
    return (
      <div id="outerbox">
        <div className="left_button">
          <h3 className="mis_title">MIS Documentation</h3>
          <button id="upload_button" onClick={this._triggerFileInput}>
            Choose File
          </button>
          <input
            type="text"
            value={this.state.fileName}
            readOnly
            placeholder="No file chosen"
            className="file-name-input"
            style={{ marginLeft: "10px", width: "200px" }}
          />
        </div>
        <div className="right_button">
          <button
            id="submit_button"
            onClick={this._handleSubmit}
            style={{ marginRight: "10px" }}
            disabled={this.state.isSubmitDisabled}
          >
            Submit
          </button>
          <button id="cancel_button" onClick={this._handleCancel}>
            Cancel
          </button>
        </div>
        {this.state.loading && <div className="loader">Loading...</div>}{" "}
        {/* Loader display */}
        <div className="outer_table">
          <input
            type="file"
            ref={(input) => (this.fileInput = input)}
            onChange={this._onFileSelected}
            style={{ display: "none" }}
            accept=".csv,.xlsx"
          />
          {this.state.tableData.length > 0 && this._renderTable()}
        </div>
      </div>
    );
  }

  private fileInput: HTMLInputElement | null = null;

  private _triggerFileInput = () => {
    if (this.fileInput) {
      this.fileInput.click();
    }
  };

  private _onFileSelected = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      const fileName = file.name;

      this.setState({ isSubmitDisabled: false, loading: true }); // Set loading to true

      const fileNameLower = fileName.toLowerCase();

      if (fileNameLower.endsWith(".csv")) {
        const fileContent = await file.text();
        this._parseCSV(fileContent);
      } else if (fileNameLower.endsWith(".xlsx")) {
        const arrayBuffer = await file.arrayBuffer();
        this._parseExcel(arrayBuffer);
      } else {
        alert("Please select a valid file type (.xlsx or .csv)");
        this.setState({ loading: false }); // Set loading to false on error
        return;
      }

      this.setState({ filePickerResult: [file], fileName });
    }
  };

  private _parseCSV = (csvContent: string) => {
    parse(csvContent, {
      header: true,
      skipEmptyLines: true,
      complete: (results: ParseResult<ITableData>) => {
        const rawData = results.data;
        this.setState({ tableData: rawData, loading: false }); // Set loading to false after parsing
      },
      error: (error: any) => {
        console.error("Error parsing CSV", error);
        this.setState({ loading: false }); // Set loading to false on error
      },
    });
  };

  private _parseExcel = (arrayBuffer: ArrayBuffer) => {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const rawData = json.slice(3); // Assuming first 3 rows are headers or irrelevant

    const data: ITableData[] = rawData.map((row: any) => ({
      "NDC Code": row[0],
      Plant: row[1],
      Dosage_form: row[2],
      Material_code: row[3],
      Description: row[4],
      Product: row[5],
      Strength: row[6],
      Pack_size: row[7],
      RMC: row[8],
      PMC: row[9],
      Consumables: row[10],
      Conversion_cost: row[11],
      Acquisition_Cost_CMO: row[12],
      Interest_on_Wc: row[13],
      COP: row[14],
      Freight_DDP_Sea: row[15],
      COGS: row[16],
      Updated_Date: this._convertExcelDate(row[17]),
      Remarks_on_Changes: row[18],
    }));

    this.setState({ tableData: data, loading: false }); // Set loading to false after parsing
  };

  private _convertExcelDate = (excelDate: number): string => {
    if (!excelDate || isNaN(excelDate)) {
      return "";
    }

    const date = new Date((excelDate - 25569) * 86400 * 1000);
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();

    return `${day}/${month}/${year}`;
  };

  private _renderTable = () => {
    const { tableData } = this.state;

    if (tableData.length === 0) return null;

    const headers = [
      "Attachment",
      "NDC Code",
      "Plant",
      "Dosage_form",
      "Material_code",
      "Description",
      "Product",
      "Strength",
      "Pack_size",
      "RMC",
      "PMC",
      "Consumables",
      "Conversion_cost",
      "Acquisition_Cost_CMO",
      "Interest_on_Wc",
      "COP",
      "Freight_DDP_Sea",
      "COGS",
      "Updated_Date",
      "Remarks_on_Changes",
    ];

    return (
      <div className="table-container">
        <table className="csv_table">
          <thead id="csv_table_head">
            <tr id="csv_header">
              {headers.map((header, index) => (
                <th key={index}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody id="csv_body">
            {tableData.map((row, rowIndex) => (
              <tr id="csv_data" key={rowIndex}>
                <td>
                  <input
                    type="file"
                    onChange={(e) =>
                      this._handleFileChange(e, row["NDC Code"])
                    }
                  />
                </td>
                {headers.slice(1).map((header, colIndex) => (
                  <td key={colIndex}>
                    {row[header as keyof ITableData]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  private _handleFileChange = (
    event: React.ChangeEvent<HTMLInputElement>,
    ndcCode: string
  ) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const attachments = [...this.state.attachments];
      attachments.push({ file: files[0], ndcCode });

      this.setState({
        attachments,
        isSubmitDisabled: attachments.length === 0,
      });
    }
  };

  private _handleCancel = () => {
    this.setState(
      {
        filePickerResult: [],
        tableData: [],
        attachments: [],
        fileName: "",
        isSubmitDisabled: false,
        loading: false, // Reset loading state on cancel
      },
      () => {
        if (this.fileInput) {
          this.fileInput.value = "";
        }
      }
    );

    localStorage.setItem("isSubmitDisabled", "false");
  };


  private _handleSubmit = async () => {
    const { tableData, attachments } = this.state;
  
    if (tableData.length === 0) {
      alert("No table data to save.");
      return;
    }
  
    this.setState({ loading: true }); // Start loading when submitting
  
    try {
      for (const row of tableData) {
        const ndcCode = row["NDC Code"];
        const plant = row["Plant"];
        const dosageForm = row["Dosage_form"];
        const material = row["Material_code"];
        const description = row["Description"];
        const product = row["Product"];
        const strength = row["Strength"].toString();
        const packsize = row["Pack_size"];
        const rmc = row["RMC"];
        const pmc = row["PMC"];
        const consumables = row["Consumables"];
        const conversionCost = row["Conversion_cost"];
        const acquisitionCostCMO = row["Acquisition_Cost_CMO"];
        const interestOnWc = row["Interest_on_Wc"];
        const cop = row["COP"];
        const freightDdpSea = row["Freight_DDP_Sea"];
        const cogs = row["COGS"];
        const updateddate = row["Updated_Date"];
        const remarksonchange = row["Remarks_on_Changes"];
  
        // Check if the item already exists in the list
        const existingItems = await sp.web.lists
          .getByTitle("MIS_Upload_File")
          .items.filter(`NDCCode eq '${ndcCode}'`)
          .top(1)
          .get();
  
        // Define the folder URL in the document library
        const folderUrl = `/sites/DevJay/MIS_Attachement/${ndcCode}`;
        const attachment = attachments.find((a) => a.ndcCode === ndcCode);
  
        // Retrieve the latest version number from the version history of the list item
        let latestVersionNumber = 1; // Default version number if no history found
  
        if (existingItems.length > 0) {
          const listItemId = existingItems[0].Id;
  
          // Get version history for the specific list item
          const versionHistory = await sp.web.lists
            .getByTitle("MIS_Upload_File")
            .items.getById(listItemId)
            .versions.get();
  
          if (versionHistory.length > 0) {
            // Get the latest version number from the version history
            latestVersionNumber = versionHistory[versionHistory.length - 1].Version;
            console.log(
              `Latest version for NDC Code '${ndcCode}' is: ${latestVersionNumber}`
            );
          }
        }
  
        // Only create the folder and upload the attachment if an attachment exists
        if (attachment) {
          try {
            // Check if the folder already exists
            await sp.web.getFolderByServerRelativeUrl(folderUrl).get();
            console.log(`Folder '${ndcCode}' already exists.`);
          } catch (e) {
            // If the folder does not exist, create it
            await sp.web.folders.add(folderUrl);
            console.log(`Folder '${ndcCode}' created.`);
  
            // Update the content type to 'Document Set'
            const folderItem = await sp.web
              .getFolderByServerRelativeUrl(folderUrl)
              .listItemAllFields.get();
  
            await sp.web.lists
              .getByTitle("MIS_Attachement")
              .items.getById(folderItem.Id)
              .update({
                ContentTypeId: "0x0120D520", // Document Set Content Type ID
              });
  
            console.log(
              `Folder '${ndcCode}' content type changed to 'Document Set'.`
            );
          }
  
          // Handle file upload into the folder
          const file = attachment.file;
          const fileExists = await sp.web
            .getFolderByServerRelativeUrl(folderUrl)
            .files.filter(`Name eq '${file.name}'`)
            .get();
  
          if (fileExists.length > 0) {
            console.log(
              `File '${file.name}' already exists in folder '${ndcCode}', uploading as a new version.`
            );
            await sp.web
              .getFolderByServerRelativeUrl(folderUrl)
              .files.add(file.name, file, true); // Overwrite file
          } else {
            console.log(`Uploading file '${file.name}' to folder '${ndcCode}'.`);
            await sp.web
              .getFolderByServerRelativeUrl(folderUrl)
              .files.add(file.name, file, false); // Add new file
          }
  
          // Update the Version_number column in the Document Set
          const folderItem = await sp.web
            .getFolderByServerRelativeUrl(folderUrl)
            .listItemAllFields.get(); // Retrieve the folder item (Document Set)
  
          await sp.web.lists
            .getByTitle("MIS_Attachement")
            .items.getById(folderItem.Id) // Use the Document Set ID
            .update({
              Version_number: latestVersionNumber, // Update with the latest version number from the list
            });
  
          console.log(
            `Version_number updated to ${latestVersionNumber} in the MIS_Attachment document library.`
          );
        } else {
          console.log(
            `No attachments found for NDC Code '${ndcCode}'. Folder will not be created or updated.`
          );
        }
  
        // Update or add item in the "MIS_Upload_File" list
        if (existingItems.length > 0) {
          // Update existing item
          const existingItem = existingItems[0];
          await sp.web.lists
            .getByTitle("MIS_Upload_File")
            .items.getById(existingItem.Id)
            .update({
              Plant: plant,
              Dosage_form: dosageForm,
              Material_code: material.toString(),
              Description: description,
              Product: product,
              Strength: strength.toString(),
              Pack_size: packsize,
              RMC: rmc,
              PMC: pmc,
              Consumables: consumables,
              Conversion_cost: conversionCost,
              Acquisition_Cost_CMO: acquisitionCostCMO,
              Interest_on_Wc: interestOnWc,
              COP: cop,
              Freight_DDP_Sea: freightDdpSea,
              COGS: cogs,
              Updated_Date: updateddate,
              Remarks_on_Changes: remarksonchange,
              Version_number: latestVersionNumber, // Update the version number in the list
            });
  
          console.log(
            `Record with NDC Code '${ndcCode}' updated in MIS_Upload_File list.`
          );
        } else {
          // Add new item
          await sp.web.lists.getByTitle("MIS_Upload_File").items.add({
            NDCCode: ndcCode,
            Plant: plant,
            Dosage_form: dosageForm,
            Material_code: material,
            Description: description,
            Product: product,
            Strength: strength.toString(),
            Pack_size: packsize,
            RMC: rmc,
            PMC: pmc,
            Consumables: consumables,
            Conversion_cost: conversionCost,
            Acquisition_Cost_CMO: acquisitionCostCMO,
            Interest_on_Wc: interestOnWc,
            COP: cop,
            Freight_DDP_Sea: freightDdpSea,
            COGS: cogs,
            Updated_Date: updateddate,
            Remarks_on_Changes: remarksonchange,
            Version_number: latestVersionNumber, // Set the version number for new items
          });
  
          console.log(
            `New record with NDC Code '${ndcCode}' created in MIS_Upload_File list.`
          );
        }
      }
  
      alert("All data has been successfully saved.");
      console.log("All data has been successfully saved.");
    } catch (error) {
      console.error("Error saving data to SharePoint", error);
      alert("Error saving data to SharePoint.");
    }
  
    this.setState({ isSubmitDisabled: true, loading: false }); 
    localStorage.setItem("isSubmitDisabled", "true");
  };
  
}
