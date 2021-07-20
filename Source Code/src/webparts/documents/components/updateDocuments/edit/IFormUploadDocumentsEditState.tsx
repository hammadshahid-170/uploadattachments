import { IUploadDocuments } from "../Models/IUploadDocuments";
import { IUploadDocumentsDataProvider } from "../sharePointDataProvider/IUploadDocumentsDataProvider";

export interface IFormUploadDocumentsEditState {
  isBusy: boolean;
  uploadDocuments: IUploadDocuments;
  messageSended: boolean;
  uploadDocumentsDataProvider:IUploadDocumentsDataProvider;
  showEdituploadDocumentsPanel:boolean;
  _goBack:VoidFunction;
  fileAttachments: any;
  isSpinnerVisible: boolean;
  showModal:boolean;
}
