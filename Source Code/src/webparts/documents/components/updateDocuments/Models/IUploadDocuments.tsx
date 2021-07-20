import { AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles';
export interface IUploadDocuments {
    ID?: number;
    name?: string;
    Name?:string;
    Email?:string;
    Comment?:string;
    Type?:string;
    gender?:any;
    FileAttachments?: AttachmentFileInfo[];
    Attachments?: boolean;
    FileNames?: string[];
    ServerRelativeUrls?: string[];
}