import * as React from 'react';
import { IFormUploadDocumentsEditState } from './IFormUploadDocumentsEditState';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton,
  MessageBar,
  MessageBarType,
  MessageBarButton,
  Panel,
  PanelType,
  PrimaryButton,
  Link,
  Dropdown
} from 'office-ui-fabric-react';
import { UploadDocumentsDataProvider } from '../sharePointDataProvider/UploadDocumentsDataProvider';
import { IUploadDocuments } from '../Models/IUploadDocuments';
import { UploadDocuments } from '../Models/UploadDocuments';
import * as jquery from 'jquery';
import { sp, ItemAddResult, ItemUpdateResult, SPHttpClient } from "@pnp/sp";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { AttachmentFile, AttachmentFiles, AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import Downloader from 'js-file-downloader';
import Modal from 'react-awesome-modal';
import FileViewer from 'react-file-viewer';
let attachments: AttachmentFileInfo[] = [];
var FullPath = window.location.href;
var arrayOfParts = FullPath.split('/');
const Sharepoint19SiteUrl = arrayOfParts.slice(0, 5).join("/");
const RestURL = window.location.origin
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
export default class FormUploadDocumentsEdit extends React.Component<{}, IFormUploadDocumentsEditState> {
  private _uploadDocumentsDataProvider: UploadDocumentsDataProvider;
  private fileInfo: HTMLInputElement;
  private delFileAtchNames: string[];
  private previewFileUrl: string = "";
  private previewFileType: string = "";
  constructor(props) {
    super(props);
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");

    this._uploadDocumentsDataProvider = new UploadDocumentsDataProvider({});
    this.state = {
      isBusy: false,
      uploadDocuments: props.state.selectedUploadDocuments,
      uploadDocumentsDataProvider: this._uploadDocumentsDataProvider,
      messageSended: false,
      showEdituploadDocumentsPanel: props.state.showEditUploadDocumentsPanel,
      _goBack: props.state._goBack,
      fileAttachments: [],
      isSpinnerVisible: false,
      showModal:false
    };
    this.delFileAtchNames = [];
    //this.state.uploadDocuments.FileAttachments = [];
  }
  componentDidMount() {
    let classObject = this;
    this._showSavedAttachments();
    // File Download function
    jquery(".downloadAttachment").click(function () {
      const fileUrl = this.getAttribute("data-url");
      new Downloader({
        url: fileUrl
      })
        .then(function () {
        })
        .catch(function (error) {
          alert('The file could not be downloaded');
        });
    });
    // file preview function 
    jquery(".previewAttachment").click(function () {
      classObject.previewFileUrl = this.getAttribute("data-src");
      classObject.previewFileType = this.getAttribute("data-ext");
      classObject.setState({ showModal: true });
    });
    // file delete function
    jquery("#fileAttachmentsRow").on("click", ".editdeleteAttachment", function (e) {
      if (window.confirm("Are you sure you want to delete ?")) {
        jquery("#browseFile").val(null);
        let { fileAttachments } = classObject.state;
        let fileName = this.id;
        if (fileAttachments.some(function (file) {
          return file.name === fileName && file.content === "This is my content";
        })) {
          classObject.delFileAtchNames.push(fileName);
        }
        fileAttachments = fileAttachments.filter(file => {
          return file.name !== fileName;
        });
        jquery(this).parent().parent().remove();
        classObject.setState({ fileAttachments: fileAttachments });
      }
      e.stopImmediatePropagation();
    });
  }
  public render(): React.ReactElement<{}> {
    const { uploadDocuments, isSpinnerVisible } = this.state;
    return (
      <div>
         {this.state.isSpinnerVisible &&
          <div className="bg-white position-absolute sploader" style={DIVSPINNER}>
            <div style={SPINNER}><Spinner className="mt-2" size={SpinnerSize.large} /></div>
          </div>
        }
         <Modal
          visible={this.state.showModal}
          width={"90%"}
          effect="fadeInDown"
          style={{ 'overflow': 'hidden' }}
        >
          <div id="filePreviewModal" className='container-fluid'>
            <div className='row'>
              <div className="col-12 p-2 bg-themeColor" style={{ background: '#ff582b' }}>
                <span className="font-weight-bold h6 text-white">File Preview</span>
                <i onClick={() => this.setState({
                  showModal:false
                })} className="ms-Icon ms-Icon--ChromeClose text-white float-right" title="Luk" style={{ cursor: 'pointer', fontSize: '16px' }} aria-hidden="true"></i>
              </div>
            </div>
            <div className="row">
              <div className="col-12 px-0" style={{ minHeight: 250, maxHeight: 300 }}>
                {this.state.showModal &&
                  <FileViewer
                    fileType={this.previewFileType}
                    filePath={this.previewFileUrl}
                  />
                }

              </div>
            </div>
            <div className="row align-items-end">
              <div className="col-12 pb-2">
                <hr className="my-1" />
                <DefaultButton style={{ float: 'right' }} onClick={() =>this.setState({
                  showModal:false
                })} >close</DefaultButton>
              </div>
            </div>
          </div>
        </Modal>
        <TextField label="Name"
          name="text" value={uploadDocuments.Name} onChanged={(event) => {
            uploadDocuments.Name = event;
            this.setState({ uploadDocuments: uploadDocuments });
          }} />
        <TextField label="Email"
          name="text" value={uploadDocuments.Email} onChanged={(event) => {
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
          selectedKey={uploadDocuments.Type}
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
            <div className="input-group">
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
          <DefaultButton text="Cancel" className="isBtnDisable mycustombutton"
            style={{ float: "right", background: 'none', borderColor: '#ff582b', color: '#ff582b', borderRadius: '5px' }}
            onClick={() => this.state._goBack()} />
          &nbsp;
          <DefaultButton text="Upload" className="isBtnDisable mycustombutton" onClick={() => this._UpdateDocuments()}
            style={{ float: "right", marginRight: "10px", background: 'none', borderColor: '#ff582b', color: '#ff582b', borderRadius: '5px' }}

          />
          &nbsp;
        </div>
      </div>
    );
  }
  ////#region show Uploaded Attachments
  private _showSavedAttachments() {
    try {
      const { uploadDocuments, fileAttachments } = this.state;
      let fileUrl: string;
      jquery("#fileAttachmentsRow").empty();
      for (let i = 0; i < uploadDocuments.ServerRelativeUrls.length; i++) {
        fileUrl = RestURL + uploadDocuments.ServerRelativeUrls[i];
        const fileExtension = uploadDocuments.FileNames[i].split(".")[1].toLowerCase();
        if (fileExtension === "jpg" || fileExtension === "jpeg" || fileExtension === "png" || fileExtension === "gif" || fileExtension === "tiff") {
          jquery("#fileAttachmentsRow").append(
            "<div class='ms-Grid-col ms-sm12 ms-md6 ms-lg6 mt-2'>" +
            "<div title='" + uploadDocuments.FileNames[i] + "' class='border border-success'>" +
            "<img class='p-1' height='50' data-action='zoom' width='70' alt='" + uploadDocuments.FileNames[i] + "' src='" + fileUrl + "' />" +
            "<button type='button' title='Delete' class='btn btn-outline-danger btn-sm rounded-0 float-right my-2 mr-2 editdeleteAttachment' style='margin-left:5px;cursor:pointer' id='" + uploadDocuments.FileNames[i] + "'><i class='fa fa-trash' aria-hidden='true'></i></button>" +
            "<button type='button' data-src='" + fileUrl + "' data-ext='" + fileExtension + "' target='_blank' title='Preview' style='margin-left:5px;cursor:pointer' class='btn btn-outline-info btn-sm rounded-0 float-right m-2 previewAttachment'><i class='fa fa-file-o' aria-hidden='true'></i></button>" +
            "<button type='button' data-url='" + fileUrl + "' title='Download' style='margin-left:5px;color:#728FCE;cursor:pointer' class='btn btn-outline-secondary btn-sm rounded-0 float-right my-2 ml-2 downloadAttachment'><i class='fa fa-long-arrow-down' aria-hidden='true'></i></button>" +
            "</div>" +
            "</div>"
          );
        }
        else if (fileExtension !== "zip" && fileExtension !== "rar" && fileExtension !== "exe") {
          jquery("#fileAttachmentsRow").append(
            "<div class='ms-Grid-col ms-sm12 ms-md6 ms-lg6 mt-2'>" +
            "<div title='" + uploadDocuments.FileNames[i] + "' class='border border-success' style='height: 52px;'>" +
            "<span class='d-inline-block text-truncate pl-2' style='padding-top: 13px; max-width: 130px;'>" + uploadDocuments.FileNames[i] + "</span>" +
            "<button type='button' title='Delete' class='btn btn-outline-danger btn-sm rounded-0 float-right my-2 mr-2 editdeleteAttachment' style='margin-left:5px;cursor:pointer' id='" + uploadDocuments.FileNames[i] + "'><i class='fa fa-trash' aria-hidden='true'></i></button>" +
            "<button type='button' data-src='" + fileUrl + "' data-ext='" + fileExtension + "' target='_blank' title='Preview' style='margin-left:5px;cursor:pointer' class='btn btn-outline-info btn-sm rounded-0 float-right m-2 previewAttachment'><i class='fa fa-file-o' aria-hidden='true'></i></button>" +
            "<button type='button' data-url='" + fileUrl + "' title='Download' class='btn btn-outline-secondary btn-sm rounded-0 float-right my-2 ml-2 downloadAttachment' style='margin-left:5px;cursor:pointer'><i class='fa fa-long-arrow-down' aria-hidden='true'></i></button>" +
            "</div>" +
            "</div>"
          );
        }
        fileAttachments.push({
          name: uploadDocuments.FileNames[i],
          content: "This is my content"
        });
      }
      console.log('showSavedAttachments', fileAttachments)
      this.setState({ fileAttachments: fileAttachments });
    } catch (error) {
      console.log('Catch Error while showSavedAttachments (in Edit File): ' + error);
    }
  }
  //#endregion
  //#region upload document function
  private _uploadDocumentfirst = (event: React.ChangeEvent<HTMLInputElement>): void => {
    try {
      let {fileAttachments } = this.state;
      const fileExtension = this.fileInfo.files[0].name.substring(this.fileInfo.files[0].name.lastIndexOf(".") + 1, this.fileInfo.files[0].name.length);
      const fileNam = this.fileInfo.files[0].name.substring(0, this.fileInfo.files[0].name.lastIndexOf(".") + 1).replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, '') + "." + fileExtension;
      const { uploadDocuments } = this.state;
      const fileSize = this.fileInfo.files[0].size / 1024 / 1024; // file size convert in MB
      const fileName = fileNam;
      if (fileSize < 10) {
        if (fileExtension === "jpg" || fileExtension === "jpeg" || fileExtension === "png" || fileExtension === "gif" || fileExtension === "tiff") {
          this.state.fileAttachments.push({
            name: fileName,
            content: this.fileInfo.files[0]
          });
          let fileUrl = URL.createObjectURL(event.target.files[0]);
          jquery("#fileAttachmentsRow").append(
            "<div class='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mt-2'>" +
            "<div title='" + fileName + "' class='border border-success'>" +
            "<img class='p-1' height='50' data-action='zoom' width='70' alt='" + fileName + "' src='" + fileUrl + "' />" +
            "<button type='button' title='Delete' style='margin-left:6px' class='btn btn-outline-danger btn-sm rounded-0 float-right m-2 editdeleteAttachment' id='" + fileName + "'><i class='fa fa-trash' aria-hidden='true'></i></button>" +
            "</div>" +
            "</div>"
          );
        }
        else
          if (!(fileExtension === "exe" || fileExtension === "zip" || fileExtension === "rar")) {
            this.state.fileAttachments.push({
              name: fileName,
              content: this.fileInfo.files[0]
            });
            jquery("#fileAttachmentsRow").append(
              "<div class='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mt-2'>" +
              "<div title='" + fileName + "' class='border border-success' style='height: 52px;'>" +
              "<span class='d-inline-block text-truncate pl-2' style='padding-top: 13px; '>" + fileName + "</span>" +
              "<button type='button' title='Delete' style='margin-left:6px' class='btn btn-outline-danger btn-sm rounded-0 float-right m-2 editdeleteAttachment' id='" + fileName + "'><i class='fa fa-trash' aria-hidden='true'></i></button>" +
              "</div>" +
              "</div>"
            );
          } else {
            jquery("#browseFile").val(null);
            window.alert("wrong format");
          }
      }
      else {
        // event.currentTarget.value = "";
        jquery("#browseFile").val(null);
      }
      this.setState({ fileAttachments: fileAttachments });
      jquery("#browseFile").val(null);
    } catch (error) {
      console.log("Catch Error while uploading document (in Create File): " + error);
    }
  }
  //#endregion
  // function for save item
  private _UpdateDocuments() {
    try {
      const { uploadDocuments } = this.state;
    uploadDocuments.FileAttachments = this.state.fileAttachments.filter(file => {
      return file.content !== "This is my content";
    });
    //this.setState({ isSpinnerVisible: true });
    this._uploadDocumentsDataProvider.updateItem(uploadDocuments, this.delFileAtchNames).then((docs: IUploadDocuments[]) => {
     // this.setState({ isSpinnerVisible: false });
      this.state._goBack();
    });
    } catch (error) {
      console.log("error while updating item",error);
    }
    
  }
}

