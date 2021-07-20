import { IUploadDocuments } from "../Models/IUploadDocuments";
import { IUploadDocumentsDataProvider } from "../sharePointDataProvider/IUploadDocumentsDataProvider";

export interface IFormUploadDocumentsCreateState {
  isBusy: boolean;
  uploadDocuments: IUploadDocuments;
  messageSended: boolean;
  isSpinnerVisible: boolean;
  uploadDocumentsDataProvider:IUploadDocumentsDataProvider;
  _goBack:VoidFunction;
  _reload:VoidFunction;
}
