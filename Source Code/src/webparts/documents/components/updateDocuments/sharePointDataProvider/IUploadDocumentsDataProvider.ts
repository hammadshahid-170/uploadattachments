import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {IUploadDocuments}  from '../Models/IUploadDocuments';
export interface IUploadDocumentsDataProvider {
    webPartContext: IWebPartContext;
    getItems(): Promise<IUploadDocuments[]>;
    createItem(itemCreated: IUploadDocuments): Promise<IUploadDocuments[]>;
    updateItem(itemUpdated: IUploadDocuments,delFileAtchNames:any): Promise<IUploadDocuments[]>;
    deleteItem(itemDeleted: IUploadDocuments): Promise<IUploadDocuments[]>;
  }
 