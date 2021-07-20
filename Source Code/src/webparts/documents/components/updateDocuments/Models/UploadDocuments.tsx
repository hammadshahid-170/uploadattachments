import {IUploadDocuments} from './IUploadDocuments';
import { AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles';

export class UploadDocuments implements IUploadDocuments {
    ID: number;
    name: string;
    value: string;
    Name:'';
    Email:'';
    Type:'';
    gender:any;
    Comment?:'';
    FileAttachments?: AttachmentFileInfo[];
    Attachments?: boolean;
    FileNames?: string[];
    ServerRelativeUrls?: string[];
}
