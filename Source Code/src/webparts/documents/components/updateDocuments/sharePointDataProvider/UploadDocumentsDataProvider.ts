import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IUploadDocuments } from '../Models/IUploadDocuments';
import { IUploadDocumentsDataProvider } from './IUploadDocumentsDataProvider';
import { sp, ItemAddResult } from "@pnp/sp";
import * as $ from 'jquery';
var FullPath = window.location.href;
var arrayOfParts = FullPath.split('/');
const Sharepoint19SiteUrl = arrayOfParts.slice(0, 4).join("/");
const BASEURL =Sharepoint19SiteUrl// "https://instantdk.sharepoint.com/sites/Test"//Sharepoint19SiteUrl;
const LIST_NAME = "Test";

export class UploadDocumentsDataProvider implements IUploadDocumentsDataProvider {
  constructor(props: {}) {
    sp.setup({
      sp: {
        headers: {
          Accept: "application/json;odata=verbose",
        },
        baseUrl: BASEURL
      },
    });

  }
  private _listuploadDocumentsUrl: string;
  private _listsUrl: string;
  private _webPartContext: IWebPartContext;
  private _uploadDocuments: IUploadDocuments[];
  public set webPartContext(value: IWebPartContext) {
    this._webPartContext = value;
    this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
  }
  public get webPartContext(): IWebPartContext {
    return this._webPartContext;
  }
  public getItems(): Promise<IUploadDocuments[]> {
    let uploadDocs: IUploadDocuments[] = [];
    // get all the items from the list in SharePoint
    return sp.web.lists.getByTitle(LIST_NAME).items.select("*","AttachmentFiles").expand("AttachmentFiles").get().then((result_docs: any[]) => {
      result_docs.forEach(element => {
        if (typeof element != 'undefined' && element) {
          let FilesUrls = [];
          let FileNames = [];
          if (element.AttachmentFiles.results.length != 0) {
            for (var i = 0; i < element.AttachmentFiles.results.length; i++) {
              FilesUrls.push(element.AttachmentFiles.results[i].ServerRelativeUrl);
              FileNames.push(element.AttachmentFiles.results[i].FileName);
            }
          }
          uploadDocs.push({
            ID: element.ID,
            Name: element.Name,
            Email: element.Email,
            Type: element.DType,
            gender: element.gender,
            Attachments: element.Attachments,
            FileNames: FileNames,
            ServerRelativeUrls: FilesUrls,
          });

        }

      });
      return uploadDocs;
    });

  }
  // get  // get all the items from the list in SharePoint
  public getItemsById(Itemid): Promise<IUploadDocuments[]> {
    try {
      let uploadDocs: IUploadDocuments[] = [];
      return sp.web.lists.getByTitle(LIST_NAME).items.getById(Itemid).select("*", "AttachmentFiles").expand("AttachmentFiles").get().then((element: any) => {
        if (typeof element != 'undefined' && element) {
          let FilesUrls = [];
          let FileNames = [];
          if (element.AttachmentFiles.results.length != 0) {
            for (var i = 0; i < element.AttachmentFiles.results.length; i++) {
              FilesUrls.push(element.AttachmentFiles.results[i].ServerRelativeUrl);
              FileNames.push(element.AttachmentFiles.results[i].FileName);
            }
          }
          uploadDocs.push({
            ID: element.ID,
            Name: element.Name,
            Email: element.Email,
            Comment:element.Comment,
            Type: element.DType,
            gender: element.gender,
            Attachments: element.Attachments,
            FileNames: FileNames,
            ServerRelativeUrls: FilesUrls,
          });
        }
        return uploadDocs;
      });
    } catch (error) {
      console.log("Error while getting item by id", error)
    }

  }
  //item create in SharePoint list
  public createItem(itemCreated: IUploadDocuments): Promise<IUploadDocuments[]> {
    let uploadDocs: IUploadDocuments[] = [];
    return sp.web.lists.getByTitle(LIST_NAME).items.add({
      Name: itemCreated.Name,
      Email: itemCreated.Email,
      DType: itemCreated.Type,
      gender: itemCreated.gender,
      Comment:itemCreated.Comment
    }).then((iar: ItemAddResult) => {
      if (itemCreated.FileAttachments.length > 0) {
        setTimeout(function () {
          iar.item.attachmentFiles.addMultiple(itemCreated.FileAttachments).then(res => {
            $(".saveLoader").hide();
            $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
            $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
            $(".ms-Panel-closeButton").trigger("click");
          }).catch(errorMsg => {

          });
        }, 2000);
      } else {
        $(".saveLoader").hide();
        $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
        $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
        $(".ms-Panel-closeButton").trigger("click");
      }
      uploadDocs.push(itemCreated);
      return uploadDocs;

    });
  }
  public updateItem(itemUpdated: IUploadDocuments,delFileAtchNames:any): Promise<IUploadDocuments[]> {
    // update an item to the list
    let uploadDocs: IUploadDocuments[] = [];
    let id = itemUpdated.ID;
    return sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      Name: itemUpdated.Name,
      Email: itemUpdated.Email,
      DType: itemUpdated.Type,
      gender: itemUpdated.gender,
      Comment:itemUpdated.Comment
    }).then((result_customers) => {
      console.log(result_customers);
      if (delFileAtchNames.length > 0) {
        setTimeout(function () {
          result_customers.item.attachmentFiles.recycleMultiple(...delFileAtchNames) 
            .then(r => {
              console.log("recycle: ", r);
              if (itemUpdated.FileAttachments.length > 0) {
                result_customers.item.attachmentFiles.addMultiple(itemUpdated.FileAttachments)
                  .then(res => {
                    console.log("Add attachment Files: ", res);
                    $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
                    $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
                    $(".ms-Panel-closeButton").trigger("click");
                  })
                  .catch(errorMsg => {
                    console.log("File Attachments Error: ", errorMsg);
                    $(".saveLoader").hide();
                    $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
                    $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
                  });
              } else {
                $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
                $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
                $(".ms-Panel-closeButton").trigger("click");
              }
            })
            .catch(error => {
              console.log("Recycle Error: ", error);
            });
        }, 3000);                   //}
      } else {
        if (itemUpdated.FileAttachments.length > 0) {
          result_customers.item.attachmentFiles.addMultiple(itemUpdated.FileAttachments)
            .then(res => {
              console.log("Add attachment Files: ", res);
              $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
              $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
              $(".ms-Panel-closeButton").trigger("click");
            })
            .catch(errorMsg => {
              console.log("File Attachments Error: ", errorMsg);
              $(".saveLoader").hide();
              $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
              $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
            });

        } else {
          $(".isBtnDisable, .ms-Panel-closeButton").removeClass("disableBtn");
          $(".deleteAttachment, .downloadAttachment, .previewAttachment, #browseFile").prop("disabled", false);
          $(".ms-Panel-closeButton").trigger("click");
        }
      }
      uploadDocs.push(itemUpdated);
      return uploadDocs;
    });
  }
  //item delete in SharePoint list
  public deleteItem(itemDeleted: IUploadDocuments): Promise<IUploadDocuments[]> {
    try {
      let id = itemDeleted.ID;
      let uploadDocs: IUploadDocuments[] = [];
      return sp.web.lists.getByTitle(LIST_NAME).items.getById(id).recycle().then((result_uploadDocs) => {
        uploadDocs.push(itemDeleted);
        return uploadDocs;
      });
    } catch (error) {
      console.log("Error while deleting Item: ", error);
    }
  }

}
function replaceNullsByEmptyString(value) {
  return (value == null) ? "" : value
}