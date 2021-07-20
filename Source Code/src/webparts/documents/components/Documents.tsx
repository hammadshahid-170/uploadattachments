import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps } from './IDocumentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {DetailsListUploadDocuments} from './updateDocuments/list/DetailsListUploadDocuments'
export default class Documents extends React.Component<IDocumentsProps, {}> {
  public render(): React.ReactElement<IDocumentsProps> {
    return(
      <DetailsListUploadDocuments/>
    );
  }
}
