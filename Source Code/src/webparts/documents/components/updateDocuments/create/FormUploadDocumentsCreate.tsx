import * as React from 'react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IFormUploadDocumentsCreateState } from './IFormUploadDocumentsCreateState';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton,
  MessageBar,
  MessageBarType,
  MessageBarButton,
  Dropdown
} from 'office-ui-fabric-react';
import { UploadDocumentsDataProvider } from '../sharePointDataProvider/UploadDocumentsDataProvider';
import { UploadDocuments } from '../Models/UploadDocuments';
import { IUploadDocuments } from '../Models/IUploadDocuments';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as jquery from 'jquery';
import { sp, ItemAddResult, ItemUpdateResult, SPHttpClient } from "@pnp/sp";
import { AttachmentFile, AttachmentFiles, AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles';
import { SPComponentLoader } from '@microsoft/sp-loader';
import "bootstrap/dist/css/bootstrap.min.css";
let attachments: AttachmentFileInfo[] = [];
var FullPath = window.location.href;
var arrayOfParts = FullPath.split('/');
const Sharepoint19SiteUrl = arrayOfParts.slice(0, 5).join("/");
const SPINNER = {
  zIndex: 9001,
  marginTop: '300px',
}
const DIVSPINNER = {
  'top':'0',
  'left': '0',
  'width': '100%',
  'height': '100%',
  zIndex: 9000,
  'opacity': 0.5
}
export default class FormUploadDocumentsCreate extends React.Component<{}, IFormUploadDocumentsCreateState> {
  private _uploadDocumentsDataProvider: UploadDocumentsDataProvider;
  private _uploadDocuments: UploadDocuments;
  private fileInfo: HTMLInputElement;
  constructor(props) {
    super(props);
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    this._uploadDocumentsDataProvider = new UploadDocumentsDataProvider({});
    this.state = {
      isBusy: false,
      uploadDocuments: new UploadDocuments(),
      uploadDocumentsDataProvider: this._uploadDocumentsDataProvider,
      messageSended: false,
      _goBack: props.state._goBack,
      _reload: props.state._reload,
      isSpinnerVisible: false
    };
    this.state.uploadDocuments.FileAttachments = [];
  }
  private _listsUrl: string;
  private _webPartContext: IWebPartContext;
  public set webPartContext(value: IWebPartContext) {
    this._webPartContext = value;
    this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
  }
  public get webPartContext(): IWebPartContext {
    return this._webPartContext;
  }
  componentDidMount() {

    // jquery function use for delete attachment from ui 
    jquery("#fileAttachmentsRow").on("click", ".deleteAttachment", function (e) {
      if (window.confirm("Are you sure you want to delete ?")) {
        jquery("#browseFile").val(null);
        const fileName = this.id;
        attachments = attachments.filter(file => {
          return file.name !== fileName;
        });
        jquery(this).parent().parent().remove();
      }
      e.stopImmediatePropagation();
    });
  }
  public render(): React.ReactElement<{}> {
    const { uploadDocuments, isSpinnerVisible } = this.state;
    return (

      <div>
        {this.state.isSpinnerVisible &&
          <div className="bg-white position-absolute" style={DIVSPINNER}>
            <div style={SPINNER}><Spinner className="mt-2" size={SpinnerSize.large} /></div>
          </div>
        }
        <TextField label="Name"
          name="text" value={uploadDocuments.Name}
          onChanged={(event) => {
            uploadDocuments.Name = event;
            this.setState({ uploadDocuments: uploadDocuments });
          }} />
        <TextField label="Email"
          name="text" value={uploadDocuments.Email}
          onChanged={(event) => {
            uploadDocuments.Email = event;
            this.setState({ uploadDocuments: uploadDocuments });
          }} />
        <Dropdown
          label='Type'
          id="type"
          onChanged={(event): void => {
            uploadDocuments.Type = event.text;
            this.setState({ uploadDocuments: uploadDocuments });
          }}
          options={
            [
              { key: 'Premium', text: 'Premium' },
              { key: 'Standard', text: 'Standard' },
              { key: 'Elite', text: 'Elite' },
            ]
          }
        />
        <div className="ms-Grid">
          <div className="ms-Grid-row" >
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style={{ 'padding': '10px 0px 0px 0px' }}>
              <label>Gender</label>
              <label>
                <input
                  type="radio"
                  name="gender"
                  value="Male"
                  style={{ marginLeft: '29px' }}
                  checked={uploadDocuments.gender === "Male"}
                  onChange={(event): void => {
                    uploadDocuments.gender = event.target.value;
                    this.setState({ uploadDocuments: uploadDocuments });
                  }}
                />{" "}
          Male{" "}
              </label>
              <label>
                <input
                  type="radio"
                  name="gender"
                  value="Female"
                  checked={uploadDocuments.gender === "Female"}
                  onChange={(event): void => {
                    uploadDocuments.gender = event.target.value;
                    this.setState({ uploadDocuments: uploadDocuments });
                  }}
                  style={{ marginLeft: '51px' }}
                />{" "}
          Female{" "}
              </label>
            </div>
          </div>
        </div>
        <TextField label="Comment" multiline rows={4}
          resizable={false}
          name="text" value={uploadDocuments.Comment} onChanged={(event) => {
            uploadDocuments.Comment = event;
            this.setState({ uploadDocuments: uploadDocuments });
          }} />
        <br></br>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <label>
              Upload Documents
            </label>
            <div className="input-group" style={{ marginTop: '10px' }}>
              <div className="custom-file">
                <input
                  type="file"
                  className="custom-file-input"
                  id="browseFile"
                  onChange={this._uploadDocumentfirst}
                  ref={(element) => { this.fileInfo = element }}
                />
                <label
                  className="custom-file-label rounded-0"
                  htmlFor="browseFile"
                  aria-describedby="inputGroupFileAddon"
                >
                   choose document 
                </label>
              </div>
            </div>
          </div>
        </div>
        <div className="ms-Grid-row custom-row align-items-center" style={{ marginTop: '10px' }} id="fileAttachmentsRow">
        </div>
        <div style={{ marginTop: "10px" }}>
          <DefaultButton text="Cancel"
            style={{ float: "right", background: 'none', borderColor: '#ff582b', color: '#ff582b', borderRadius: '5px' }}
            className="isBtnDisable mycustombutton" onClick={() => this.state._goBack()} />
          &nbsp;
          <DefaultButton text="Upload" className="isBtnDisable mycustombutton"
            onClick={() => this._SaveDocuments()}
            style={{ float: "right", marginRight: "10px", background: 'none', borderColor: '#ff582b', color: '#ff582b', borderRadius: '5px' }}
          />
          &nbsp;
        </div>
      </div>
    );
  }
  //#region upload document function
  private _uploadDocumentfirst = (event: React.ChangeEvent<HTMLInputElement>): void => {
    try {
      const fileExtension = this.fileInfo.files[0].name.substring(this.fileInfo.files[0].name.lastIndexOf(".") + 1, this.fileInfo.files[0].name.length);
      const fileNam = this.fileInfo.files[0].name.substring(0, this.fileInfo.files[0].name.lastIndexOf(".") + 1).replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, '') + "." + fileExtension;
      const { uploadDocuments } = this.state;
      const fileSize = this.fileInfo.files[0].size / 1024 / 1024; // file size convert in MB
      const fileName = fileNam;
      if (fileExtension === "jpg" || fileExtension === "jpeg" || fileExtension === "png" || fileExtension === "gif" || fileExtension === "tiff") {
        attachments.push({
          name: fileName,
          content: this.fileInfo.files[0]
        });
        let fileUrl = URL.createObjectURL(event.target.files[0]);
        jquery("#fileAttachmentsRow").append(
          "<div class='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mt-2'>" +
          "<div title='" + fileName + "' class='border border-success'>" +
          "<img class='p-1' height='50' data-action='zoom' width='70' alt='" + fileName + "' src='" + fileUrl + "' />" +
          "<button type='button' title='Slet' class='btn btn-outline-danger btn-sm rounded-0 float-right m-2 deleteAttachment' id='" + fileName + "' style='margin-left:5px;cursor:pointer'><i class='fa fa-trash' aria-hidden='true' ></i></button>" +
          "</div>" +
          "</div>"
        );
      }
      else
        if (!(fileExtension === "exe" || fileExtension === "zip" || fileExtension === "rar")) {
          attachments.push({
            name: fileName,
            content: this.fileInfo.files[0]
          });
          jquery("#fileAttachmentsRow").append(
            "<div class='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mt-2'>" +
            "<div title='" + fileName + "' class='border border-success' style='height: 52px;'>" +
            "<span class='d-inline-block text-truncate pl-2' style='padding-top: 13px; '>" + fileName + "</span>" +
            "<button type='button' title='Slet'  class='btn btn-outline-danger btn-sm rounded-0 float-right m-2 deleteAttachment' id='" + fileName + "' style='margin-left:5px;cursor:pointer'><i class='fa fa-trash' aria-hidden='true'> </i></button>" +
            "</div>" +
            "</div>"
          );
        } else {
          jquery("#browseFile").val(null);
          window.alert("wrong format");
        }


      jquery("#browseFile").val(null);
    } catch (error) {
      console.log("Catch Error while uploading document (in Create File): " + error);
    }
  }
  //#endregion
  // function for save item
  private _SaveDocuments() {
    try {
      const { uploadDocuments } = this.state;
    uploadDocuments.FileAttachments = attachments;
    attachments = [];
    this.setState({ isSpinnerVisible: true });
    this._uploadDocumentsDataProvider.createItem(uploadDocuments).then((docs: IUploadDocuments[]) => {
     this.setState({isSpinnerVisible:false});
      this.state._reload();
    });
    this.state._goBack();
    } catch (error) {
      console.log("error while creating item",error);
    }
    
  }
}