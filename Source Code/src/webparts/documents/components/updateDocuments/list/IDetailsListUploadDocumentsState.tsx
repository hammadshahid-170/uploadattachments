import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IUploadDocuments } from "../Models/IUploadDocuments";
export interface IDetailsListUploadDocumentsState {
    columns: IColumn[];
    items: IUploadDocuments[];
    selectionDetails: string;
    selectedUploadDocuments: IUploadDocuments;
    showEditUploadDocumentsPanel:boolean;
    showModal:boolean;
    _goBack:VoidFunction;
    _reloadList?:VoidFunction;
}